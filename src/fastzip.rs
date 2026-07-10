//! A minimal, in-memory ZIP writer tuned for jetxl's workload.
//!
//! # Why this exists
//!
//! jetxl produces every archive part as a `Vec<u8>` that already lives in
//! memory, then hands it to a general-purpose zip library. General crates
//! (mtzip, zip-rs) drive their DEFLATE encoder through the `Read` trait, which
//! forces the compressor to repeatedly copy the already-in-memory bytes into a
//! temporary staging buffer before compressing them. That copy is pure
//! overhead for our case: the data is a contiguous slice we could hand to the
//! compressor directly. Removing that indirection is a measurable win with
//! *no* change to the output bytes or compression ratio (the DEFLATE stream is
//! byte-for-byte what miniz_oxide would produce either way).
//!
//! This writer compresses each part with `miniz_oxide` directly over its slice,
//! computes CRC32 with `crc32fast`, and emits the ZIP structure (local headers,
//! central directory, EOCD) by hand. The output is a standard DEFLATE-based ZIP
//! and therefore a spec-valid `.xlsx` that Excel/LibreOffice/openpyxl open
//! normally.
//!
//! # Format scope
//!
//! Only what OOXML needs: DEFLATE and Stored entries, no directories (OOXML
//! paths are implicit), no Zip64 (worksheets are far under the 4 GiB per-entry
//! limit), no encryption. Entries carry a fixed MS-DOS timestamp for
//! reproducible output. This is deliberately not a general ZIP library.
//!
//! # Parallelism
//!
//! Parts are independent, so compression is done with rayon's `par_iter`. On a
//! multi-core machine this overlaps the DEFLATE work across all parts; on a
//! single core it degrades to sequential with negligible overhead. Assembly of
//! the final byte stream is serial (it must be, ordering matters) but is just
//! header formatting plus `extend_from_slice` of already-compressed buffers.

/// Compression selector. Mirrors the small slice of mtzip's `CompressionLevel`
/// API that jetxl used, so call sites port over unchanged.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CompressionLevel(u8);

impl CompressionLevel {
    /// Fastest useful DEFLATE (miniz_oxide level 1). Matches the old default.
    #[inline]
    pub const fn fast() -> Self {
        CompressionLevel(1)
    }

    /// No compression: the part is stored verbatim.
    #[allow(dead_code)]  // API completeness (mirrors mtzip's level set)
    #[inline]
    pub const fn none() -> Self {
        CompressionLevel(0)
    }

    /// Best ratio (miniz_oxide level 9), slowest.
    #[allow(dead_code)]  // API completeness (mirrors mtzip's level set)
    #[inline]
    pub const fn best() -> Self {
        CompressionLevel(9)
    }

    #[inline]
    fn level(self) -> u8 {
        self.0
    }
}

/// ZIP compression method identifiers (APPNOTE 4.4.5).
const METHOD_STORE: u16 = 0;
const METHOD_DEFLATE: u16 = 8;

/// A single pending entry: either raw bytes to be compressed at write time, or
/// an already-compressed part supplied by the caller. Deferring raw compression
/// to `write` lets it be parallelized across all entries at once; the
/// already-compressed variant lets the caller pipeline compression with its own
/// upstream work (e.g. XML generation) so the two phases overlap.
enum PendingEntry {
    Raw {
        name: String,
        data: Vec<u8>,
        level: CompressionLevel,
    },
    Prepared(CompressedEntry),
}

/// An already-compressed archive part, produced by [`compress_part`]. Handing
/// these to [`ZipArchive::add_prepared`] lets the caller compress a part inside
/// its own parallel pipeline stage instead of waiting for the ZIP writer's
/// separate compression phase.
pub struct PreparedPart(CompressedEntry);

/// Compress one part now, off to the side, so the caller can overlap it with
/// other work. Same DEFLATE/store logic and byte output as the writer's own
/// compression path; the result is inserted verbatim by `add_prepared`.
pub fn compress_part(name: String, data: Vec<u8>, level: CompressionLevel) -> PreparedPart {
    PreparedPart(compress_raw(name, data, level))
}

/// Builder handle returned by `add_file_from_memory`, mirroring mtzip's
/// fluent API (`.compression_level(..).done()`). Dropping without `done()`
/// still commits the entry, but call sites always call `done()`.
pub struct FileBuilder<'a> {
    archive: &'a mut ZipArchive,
    index: usize,
}

impl<'a> FileBuilder<'a> {
    /// Set the compression level for this entry.
    #[inline]
    pub fn compression_level(self, level: CompressionLevel) -> Self {
        if let PendingEntry::Raw { level: l, .. } = &mut self.archive.entries[self.index] {
            *l = level;
        }
        self
    }

    /// Commit the entry. No-op beyond consuming the builder; present for API
    /// parity with mtzip so existing call sites compile unchanged.
    #[inline]
    pub fn done(self) {}
}

/// An in-memory ZIP archive builder. Drop-in for the subset of mtzip's API
/// that jetxl used.
pub struct ZipArchive {
    entries: Vec<PendingEntry>,
}

impl ZipArchive {
    #[inline]
    pub fn new() -> Self {
        ZipArchive {
            entries: Vec::new(),
        }
    }

    /// Queue a part for inclusion. The payload is taken by value; compression
    /// happens later in `write`. Returns a builder for setting the level.
    #[inline]
    pub fn add_file_from_memory(&mut self, data: Vec<u8>, name: String) -> FileBuilder<'_> {
        let index = self.entries.len();
        self.entries.push(PendingEntry::Raw {
            name,
            data,
            level: CompressionLevel::fast(),
        });
        FileBuilder {
            archive: self,
            index,
        }
    }

    /// Insert an already-compressed part (from [`compress_part`]) at the current
    /// position. Lets the caller pipeline compression with upstream work; the
    /// writer skips re-compressing it. Order is preserved exactly like
    /// `add_file_from_memory`, so mixing prepared and raw parts is fine.
    #[inline]
    pub fn add_prepared(&mut self, part: PreparedPart) {
        self.entries.push(PendingEntry::Prepared(part.0));
    }

    /// Compress all parts and write the complete ZIP to `writer`.
    ///
    /// Compression is parallel across parts; assembly is serial. The output
    /// buffer is pre-sized from the uncompressed total (an over-estimate, since
    /// compression shrinks it), so the serial assembly never reallocates.
    /// Compress and assemble the archive into a freshly allocated `Vec`,
    /// pre-sized from the exact compressed total so the buffer is allocated once
    /// and never reallocates during assembly. Preferred over `write` for the
    /// in-memory (`to_bytes`) path.
    pub fn write_to_vec(&mut self) -> Vec<u8> {
        let entries = std::mem::take(&mut self.entries);
        let compressed: Vec<CompressedEntry> = par_compress(entries);

        // Exact final size: per entry, 30-byte local header + name + data, plus
        // 46-byte central header + name, plus the 22-byte EOCD.
        let total: usize = compressed
            .iter()
            .map(|c| 30 + c.name.len() + c.compressed.len() + 46 + c.name.len())
            .sum::<usize>()
            + 22;

        let mut out: Vec<u8> = Vec::with_capacity(total);
        let mut central_directory: Vec<u8> = Vec::with_capacity(compressed.len() * 80);
        let mut offsets: Vec<u32> = Vec::with_capacity(compressed.len());

        for entry in &compressed {
            offsets.push(out.len() as u32);
            write_local_header(&mut out, entry);
            out.extend_from_slice(&entry.compressed);
        }
        let cd_start = out.len() as u32;
        for (entry, &off) in compressed.iter().zip(offsets.iter()) {
            write_central_header(&mut central_directory, entry, off);
        }
        let cd_size = central_directory.len() as u32;
        out.extend_from_slice(&central_directory);
        write_eocd(&mut out, compressed.len() as u16, cd_start, cd_size);
        out
    }
}

impl Default for ZipArchive {
    fn default() -> Self {
        Self::new()
    }
}

/// A part after compression: retains name, CRC, sizes, method, and the
/// compressed bytes needed to emit both headers.
struct CompressedEntry {
    name: String,
    crc: u32,
    compressed: Vec<u8>,
    uncompressed_size: u32,
    method: u16,
}

/// Compress raw bytes into a `CompressedEntry`. `Store` when the level is
/// `none()` or when DEFLATE fails to shrink the data (storing avoids paying for
/// a larger payload). CRC32 is over the *uncompressed* bytes, per spec.
fn compress_raw(name: String, data: Vec<u8>, level: CompressionLevel) -> CompressedEntry {
    let crc = crc32fast::hash(&data);
    let uncompressed_size = data.len() as u32;

    if level.level() == 0 {
        return CompressedEntry { name, crc, compressed: data, uncompressed_size, method: METHOD_STORE };
    }

    // Compress directly over the in-memory slice: no Read-trait staging buffer.
    let deflated = deflate_backend(&data, level.level());

    // If DEFLATE didn't help (tiny or incompressible parts), store verbatim so
    // the archive never grows a part.
    if deflated.len() >= data.len() {
        CompressedEntry { name, crc, compressed: data, uncompressed_size, method: METHOD_STORE }
    } else {
        CompressedEntry { name, crc, compressed: deflated, uncompressed_size, method: METHOD_DEFLATE }
    }
}

/// Resolve a pending entry to a compressed one. Raw entries are compressed;
/// already-prepared entries pass through untouched.
fn compress_one(entry: PendingEntry) -> CompressedEntry {
    match entry {
        PendingEntry::Raw { name, data, level } => compress_raw(name, data, level),
        PendingEntry::Prepared(c) => c,
    }
}

/// Produce a raw DEFLATE stream from `data` at the given level.
///
/// Two interchangeable backends, selected at compile time. Both emit standard
/// raw DEFLATE (no zlib/gzip wrapper), so the ZIP entry is identical in format
/// regardless of which is used; only speed and the exact byte output differ.
#[cfg(not(feature = "libdeflate"))]
#[inline]
fn deflate_backend(data: &[u8], level: u8) -> Vec<u8> {
    // Default: pure-Rust miniz_oxide, no C toolchain required.
    miniz_oxide::deflate::compress_to_vec(data, level)
}

/// libdeflate backend (opt-in via the `libdeflate` feature).
///
/// libdeflate's level scale is 1..=12; jetxl's levels are 1 (fast), 9 (best).
/// We map jetxl's 1 -> libdeflate 1 and clamp everything else into range. The
/// output is raw DEFLATE via `deflate_compress`, matching the miniz path.
/// `deflate_compress_bound` gives the worst-case size so the destination buffer
/// is allocated once.
#[cfg(feature = "libdeflate")]
#[inline]
fn deflate_backend(data: &[u8], level: u8) -> Vec<u8> {
    use libdeflater::{CompressionLvl, Compressor};

    let lvl_num = level.clamp(1, 12) as i32;
    let lvl = CompressionLvl::new(lvl_num).unwrap_or_else(|_| CompressionLvl::fastest());
    let mut compressor = Compressor::new(lvl);

    let bound = compressor.deflate_compress_bound(data.len());
    let mut out = vec![0u8; bound];
    match compressor.deflate_compress(data, &mut out) {
        Ok(n) => {
            out.truncate(n);
            out
        }
        // On the vanishingly unlikely failure path, fall back to miniz so a
        // write never fails for a compression-internal reason.
        Err(_) => miniz_oxide::deflate::compress_to_vec(data, level),
    }
}

/// Compress all entries in parallel, preserving input order (order is
/// significant for the archive layout).
fn par_compress(entries: Vec<PendingEntry>) -> Vec<CompressedEntry> {
    use rayon::prelude::*;
    entries.into_par_iter().map(compress_one).collect()
}

/// Fixed MS-DOS date/time for reproducible archives (2020-01-01 00:00:00).
const DOS_TIME: u16 = 0;
const DOS_DATE: u16 = 0x0021;

/// Write a local file header (APPNOTE 4.3.7) followed by nothing; caller
/// appends the compressed data.
fn write_local_header(out: &mut Vec<u8>, e: &CompressedEntry) {
    out.extend_from_slice(&0x0403_4b50u32.to_le_bytes()); // local file header signature "PK\x03\x04"
    out.extend_from_slice(&20u16.to_le_bytes()); // version needed to extract (2.0)
    out.extend_from_slice(&0u16.to_le_bytes()); // general purpose bit flag
    out.extend_from_slice(&e.method.to_le_bytes()); // compression method
    out.extend_from_slice(&DOS_TIME.to_le_bytes()); // last mod file time
    out.extend_from_slice(&DOS_DATE.to_le_bytes()); // last mod file date
    out.extend_from_slice(&e.crc.to_le_bytes()); // crc-32
    out.extend_from_slice(&(e.compressed.len() as u32).to_le_bytes()); // compressed size
    out.extend_from_slice(&e.uncompressed_size.to_le_bytes()); // uncompressed size
    out.extend_from_slice(&(e.name.len() as u16).to_le_bytes()); // file name length
    out.extend_from_slice(&0u16.to_le_bytes()); // extra field length
    out.extend_from_slice(e.name.as_bytes()); // file name
}

/// Write a central directory file header (APPNOTE 4.3.12) for one entry.
fn write_central_header(cd: &mut Vec<u8>, e: &CompressedEntry, local_offset: u32) {
    cd.extend_from_slice(&0x0201_4b50u32.to_le_bytes()); // central file header signature "PK\x01\x02"
    cd.extend_from_slice(&20u16.to_le_bytes()); // version made by
    cd.extend_from_slice(&20u16.to_le_bytes()); // version needed to extract
    cd.extend_from_slice(&0u16.to_le_bytes()); // general purpose bit flag
    cd.extend_from_slice(&e.method.to_le_bytes()); // compression method
    cd.extend_from_slice(&DOS_TIME.to_le_bytes()); // last mod file time
    cd.extend_from_slice(&DOS_DATE.to_le_bytes()); // last mod file date
    cd.extend_from_slice(&e.crc.to_le_bytes()); // crc-32
    cd.extend_from_slice(&(e.compressed.len() as u32).to_le_bytes()); // compressed size
    cd.extend_from_slice(&e.uncompressed_size.to_le_bytes()); // uncompressed size
    cd.extend_from_slice(&(e.name.len() as u16).to_le_bytes()); // file name length
    cd.extend_from_slice(&0u16.to_le_bytes()); // extra field length
    cd.extend_from_slice(&0u16.to_le_bytes()); // file comment length
    cd.extend_from_slice(&0u16.to_le_bytes()); // disk number start
    cd.extend_from_slice(&0u16.to_le_bytes()); // internal file attributes
    cd.extend_from_slice(&0u32.to_le_bytes()); // external file attributes
    cd.extend_from_slice(&local_offset.to_le_bytes()); // relative offset of local header
    cd.extend_from_slice(e.name.as_bytes()); // file name
}

/// Write the End Of Central Directory record (APPNOTE 4.3.16).
fn write_eocd(out: &mut Vec<u8>, count: u16, cd_start: u32, cd_size: u32) {
    out.extend_from_slice(&0x0605_4b50u32.to_le_bytes()); // EOCD signature "PK\x05\x06"
    out.extend_from_slice(&0u16.to_le_bytes()); // number of this disk
    out.extend_from_slice(&0u16.to_le_bytes()); // disk with start of central directory
    out.extend_from_slice(&count.to_le_bytes()); // entries on this disk
    out.extend_from_slice(&count.to_le_bytes()); // total entries
    out.extend_from_slice(&cd_size.to_le_bytes()); // size of central directory
    out.extend_from_slice(&cd_start.to_le_bytes()); // offset of central directory
    out.extend_from_slice(&0u16.to_le_bytes()); // .zip file comment length
}