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
