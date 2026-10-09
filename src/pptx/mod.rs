//! PPTX (PowerPoint) presentation parser.
//!
//! This module provides parsing for Microsoft PowerPoint presentations in the
//! Office Open XML (.pptx) format.

mod bullets;
mod parser;
#[cfg(feature = "raster")]
mod raster;

pub use parser::PptxParser;
