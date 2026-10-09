//! Slide rasterization (feature `raster`): a presentation slide painted to pixels, for a reader
//! that needs what a slide *looks like* — the shapes, connectors and fills that carry a diagram
//! slide's structure and have no text of their own.
//!
//! What could not be painted is counted in [`SlideRasterGaps`] rather than silently left out,
//! so a consumer can tell a slide drawn in full from one with holes.

/// How to rasterize a slide.
#[derive(Debug, Clone, PartialEq)]
pub struct SlideRasterOptions {
    /// Resolution in dots per inch; a slide point is `dpi / 72` pixels. Default 150.
    pub dpi: f32,
}

impl Default for SlideRasterOptions {
    fn default() -> Self {
        Self { dpi: 150.0 }
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
    /// Text runs not painted.
    pub text_runs: u32,
    /// Charts not drawn.
    pub charts: u32,
    /// Tables, diagrams (SmartArt) and other graphic frames not drawn.
    pub graphic_frames: u32,
    /// Fills drawn as a stand-in: a gradient or pattern painted in one of its colors.
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
