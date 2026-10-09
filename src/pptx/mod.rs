//! PPTX (PowerPoint) presentation parser.
//!
//! This module provides parsing for Microsoft PowerPoint presentations in the
//! Office Open XML (.pptx) format.

mod bullets;
mod parser;
#[cfg(feature = "raster")]
mod raster;
#[cfg(feature = "raster")]
mod raster_text;

pub use parser::PptxParser;

/// The slide-raster tests' presentation builders, for tests elsewhere in the crate that need
/// a presentation (the C ABI's render entry point).
#[cfg(all(test, feature = "raster"))]
pub(crate) use raster::tests as raster_fixtures;
