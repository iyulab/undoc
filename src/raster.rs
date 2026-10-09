//! Slide rasterization (feature `raster`): a presentation slide painted to pixels, for a reader
//! that needs what a slide *looks like* — the shapes, connectors and fills that carry a diagram
//! slide's structure and have no text of their own.
//!
//! What could not be painted is counted in [`SlideRasterGaps`] rather than silently left out,
//! so a consumer can tell a slide drawn in full from one with holes.

/// How to rasterize a slide.
///
/// Text is drawn in the faces in [`fonts`](Self::fonts), then those in
/// [`font_dirs`](Self::font_dirs), then — with [`system_fonts`](Self::system_fonts) — the
/// system's font directories. No face is bundled: a host without fonts (a minimal container)
/// draws no text unless faces are passed in, and reports the runs in
/// [`SlideRasterGaps::text_runs`].
#[derive(Debug, Clone, PartialEq)]
pub struct SlideRasterOptions {
    /// Resolution in dots per inch; a slide point is `dpi / 72` pixels. Default 150.
    pub dpi: f32,
    /// Font files (TrueType, OpenType, or collections of them) to draw text in.
    pub fonts: Vec<Vec<u8>>,
    /// Directories searched, with their subdirectories, for font files.
    pub font_dirs: Vec<std::path::PathBuf>,
    /// Whether the system's font directories are searched too. Default `true`.
    pub system_fonts: bool,
}

impl Default for SlideRasterOptions {
    fn default() -> Self {
        Self {
            dpi: 150.0,
            fonts: Vec::new(),
            font_dirs: Vec::new(),
            system_fonts: true,
        }
    }
}

/// What a rasterized slide could not show, by kind. All zero means everything the slide asks
/// for was painted.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SlideRasterGaps {
    /// Shapes not drawn: a geometry this does not read (custom geometry, an unknown preset).
    pub shapes: u32,
    /// Pictures not painted.
    pub images: u32,
    /// Text runs not painted, or painted only in part: no face has their characters, their
    /// script needs shaping, or the text is vertical.
    pub text_runs: u32,
    /// Charts not drawn.
    pub charts: u32,
    /// Graphic frames not drawn: embedded objects, and SmartArt PowerPoint left no drawing for.
    /// Tables and SmartArt are drawn.
    pub graphic_frames: u32,
    /// Fills drawn as a stand-in: a pattern in its foreground color, a rectangular or
    /// shape-following gradient as a radial one, a tiled picture stretched, a gradient line in
    /// its first color.
    pub approximated_fills: u32,
}

impl SlideRasterGaps {
    /// Whether anything the slide asked for was left unpainted or approximated.
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

/// A rasterized slide: `width × height` pixels, opaque RGBA, row by row from the top-left.
#[derive(Debug, Clone)]
pub struct RasteredSlide {
    pub width: u32,
    pub height: u32,
    pub rgba: Vec<u8>,
    pub gaps: SlideRasterGaps,
    /// Text runs drawn in a face standing in for the one they ask for (it is not installed, or
    /// lacks their characters): readable, but not the slide's own typeface. Not a gap.
    pub substituted_text_runs: u32,
}

impl RasteredSlide {
    /// The slide as a PNG.
    pub fn to_png(&self) -> Vec<u8> {
        let size = tiny_skia::IntSize::from_wh(self.width, self.height)
            .expect("a rastered slide has a non-zero size");
        tiny_skia::Pixmap::from_vec(self.rgba.clone(), size)
            .expect("an RGBA buffer of width * height pixels")
            .encode_png()
            .expect("an opaque pixmap encodes as PNG")
    }
}
