//! undoc CLI - Microsoft Office document extraction tool
//!
//! A command-line tool for extracting content from DOCX, XLSX, and PPTX files.

mod update;
mod writer;

use clap::{Parser, Subcommand, ValueEnum};
use colored::*;
use indicatif::{ProgressBar, ProgressStyle};
use std::fs;
use std::io::{self, Write};
use std::path::PathBuf;
use undoc::render::{CleanupPreset, HeadingConfig, JsonFormat, RenderOptions, TableFallback};
use writer::{MultiFormatWriter, OutputFormat, StreamingWriter};

/// Microsoft Office document extraction to Markdown, text, and JSON
#[derive(Parser)]
#[command(
    name = "undoc",
    author = "iyulab",
    version,
    about = "Extract content from Office documents",
    long_about = "undoc - High-performance Microsoft Office document extraction tool.\n\n\
                  Converts DOCX, XLSX, and PPTX files to Markdown, plain text, or JSON.\n\n\
                  Usage:\n  \
                  undoc <file>              Extract all formats to output directory\n  \
                  undoc <file> <output>     Extract to specified directory\n  \
                  undoc md <file>           Convert to Markdown only"
)]
struct Cli {
    #[command(subcommand)]
    command: Option<Commands>,

    /// Input file path (for default conversion)
    #[arg(global = false)]
    input: Option<PathBuf>,

    /// Output directory (for default conversion)
    #[arg(global = false)]
    output: Option<PathBuf>,

    /// Apply text cleanup preset
    #[arg(long, global = true)]
    cleanup: Option<CleanupMode>,
}

#[derive(Subcommand)]
enum Commands {
    /// Convert a document (Markdown output by default; use --all or --formats for more)
    Convert {
        /// Input file path
        input: PathBuf,

        /// Output directory (default: <filename>_output)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Apply text cleanup
        #[arg(long)]
        cleanup: Option<CleanupMode>,

        /// Output formats to produce (comma-separated: md,txt,json; default: md)
        #[arg(long, value_delimiter = ',', default_value = "md")]
        formats: Vec<String>,

        /// Produce all output formats (md + txt + json)
        #[arg(long)]
        all: bool,

        /// Skip image extraction (images are extracted by default)
        #[arg(long)]
        no_images: bool,

        /// Suppress progress output
        #[arg(short, long)]
        quiet: bool,

        /// Emit `\n\n---\n\n` for hard page breaks (default: off — markdown has
        /// no page concept).
        #[arg(long)]
        emit_page_breaks: bool,

        /// Include DOCX section headers/footers as blockquoted lines around
        /// the body (default: off — they are typically page-chrome noise).
        #[arg(long)]
        include_headers_footers: bool,

        /// Shortcut: enable both `--emit-page-breaks` and
        /// `--include-headers-footers` (i.e. `RenderOptions::lossless()`).
        #[arg(long)]
        lossless: bool,

        /// Insert HTML section boundary markers (<!-- slide N: Name -->, <!-- sheet N: Name -->).
        /// No effect on DOCX documents.
        #[arg(long)]
        section_markers: bool,
    },

    /// Convert a document to Markdown
    #[command(visible_alias = "md")]
    Markdown {
        /// Input file path
        input: PathBuf,

        /// Output file path (default: stdout)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Include YAML frontmatter with metadata
        #[arg(short, long)]
        frontmatter: bool,

        /// Table rendering mode
        #[arg(long, default_value = "markdown")]
        table_mode: TableMode,

        /// Apply text cleanup
        #[arg(long)]
        cleanup: Option<CleanupMode>,

        /// Maximum heading level (1-6, default: 4)
        #[arg(long, default_value = "4")]
        max_heading: u8,

        /// Emit `\n\n---\n\n` for hard page breaks (default: off — markdown has
        /// no page concept).
        #[arg(long)]
        emit_page_breaks: bool,

        /// Include DOCX section headers/footers as blockquoted lines around
        /// the body (default: off — they are typically page-chrome noise).
        #[arg(long)]
        include_headers_footers: bool,

        /// Shortcut: enable both `--emit-page-breaks` and
        /// `--include-headers-footers` (i.e. `RenderOptions::lossless()`).
        #[arg(long)]
        lossless: bool,

        /// Insert HTML section boundary markers (<!-- slide N: Name -->, <!-- sheet N: Name -->).
        /// No effect on DOCX documents.
        #[arg(long)]
        section_markers: bool,
    },

    /// Convert a document to plain text
    Text {
        /// Input file path
        input: PathBuf,

        /// Output file path (default: stdout)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Apply text cleanup
        #[arg(long)]
        cleanup: Option<CleanupMode>,
    },

    /// Convert a document to JSON
    Json {
        /// Input file path
        input: PathBuf,

        /// Output file path (default: stdout)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Output compact JSON (no indentation)
        #[arg(long)]
        compact: bool,
    },

    /// Show document information and metadata
    Info {
        /// Input file path
        input: PathBuf,
    },

    /// Extract resources (images, media) from a document
    Extract {
        /// Input file path
        input: PathBuf,

        /// Output directory for resources
        #[arg(short, long, default_value = ".")]
        output: PathBuf,
    },

    /// Update undoc to the latest version
    Update {
        /// Check only, don't install
        #[arg(long)]
        check: bool,

        /// Force update even if on latest version
        #[arg(long)]
        force: bool,
    },

    /// Show version information
    Version,
}

/// Table rendering mode
#[derive(Clone, ValueEnum)]
enum TableMode {
    /// Standard Markdown tables
    Markdown,
    /// HTML tables (for complex layouts)
    Html,
    /// ASCII art tables
    Ascii,
}

impl From<TableMode> for TableFallback {
    fn from(mode: TableMode) -> Self {
        match mode {
            TableMode::Markdown => TableFallback::Markdown,
            TableMode::Html => TableFallback::Html,
            TableMode::Ascii => TableFallback::Ascii,
        }
    }
}

/// Cleanup mode
#[derive(Clone, ValueEnum)]
enum CleanupMode {
    /// No cleanup
    None,
    /// Minimal cleanup
    Minimal,
    /// Standard cleanup (default)
    Standard,
    /// Aggressive cleanup
    Aggressive,
}

impl CleanupMode {
    fn to_preset(self) -> Option<CleanupPreset> {
        match self {
            Self::None => None,
            Self::Minimal => Some(CleanupPreset::Minimal),
            Self::Standard => Some(CleanupPreset::Default),
            Self::Aggressive => Some(CleanupPreset::Aggressive),
        }
    }
}

/// Check if we should perform background update check.
/// Skip for update/version commands to avoid redundant checks.
fn should_check_update(cli: &Cli) -> bool {
    !matches!(
        &cli.command,
        Some(Commands::Update { .. }) | Some(Commands::Version)
    )
}

fn main() {
    let cli = Cli::parse();

    // Start background update check (except for update/version commands)
    let update_rx = if should_check_update(&cli) {
        Some(update::check_update_async())
    } else {
        None
    };

    // Run the main command
    let result = run(cli);

    // Check for update result and show notification if available
    if let Some(rx) = update_rx {
        if let Some(update_result) = update::try_get_update_result(&rx) {
            update::print_update_notification(&update_result);
        }
    }

    // Handle errors
    if let Err(e) = result {
        eprintln!("{}: {}", "Error".red().bold(), e);
        std::process::exit(1);
    }
}

fn run(cli: Cli) -> Result<(), Box<dyn std::error::Error>> {
    // Handle default command (undoc <file> [output])
    if cli.command.is_none() {
        if let Some(input) = cli.input {
            return run_convert(ConvertParams {
                input: &input,
                output: cli.output.as_ref(),
                cleanup: cli.cleanup,
                formats: &[OutputFormat::Markdown],
                no_images: false,
                quiet: false,
                emit_page_breaks: false,
                include_headers_footers: false,
                lossless: false,
                section_markers: false,
            });
        } else {
            // No input provided, show help
            use clap::CommandFactory;
            Cli::command().print_help()?;
            return Ok(());
        }
    }

    match cli.command.unwrap() {
        Commands::Convert {
            input,
            output,
            cleanup,
            formats,
            all,
            no_images,
            quiet,
            emit_page_breaks,
            include_headers_footers,
            lossless,
            section_markers,
        } => {
            let resolved: Vec<OutputFormat> = if all {
                vec![
                    OutputFormat::Markdown,
                    OutputFormat::Text,
                    OutputFormat::Json,
                ]
            } else {
                formats
                    .iter()
                    .filter_map(|s| OutputFormat::parse(s))
                    .collect()
            };
            let resolved = if resolved.is_empty() {
                vec![OutputFormat::Markdown]
            } else {
                resolved
            };
            run_convert(ConvertParams {
                input: &input,
                output: output.as_ref(),
                cleanup,
                formats: &resolved,
                no_images,
                quiet,
                emit_page_breaks,
                include_headers_footers,
                lossless,
                section_markers,
            })?;
        }

        Commands::Markdown {
            input,
            output,
            frontmatter,
            table_mode,
            cleanup,
            max_heading,
            emit_page_breaks,
            include_headers_footers,
            lossless,
            section_markers,
        } => {
            let pb = create_spinner("Parsing document...");

            let doc = undoc::parse_file(&input)?;
            pb.set_message("Rendering to Markdown...");

            let heading_config = HeadingConfig::default().with_default_style_mapping();
            let base = if lossless {
                RenderOptions::lossless()
            } else {
                RenderOptions::new()
                    .with_emit_page_breaks(emit_page_breaks)
                    .with_include_headers_footers(include_headers_footers)
            };
            let mut options = base
                .with_frontmatter(frontmatter)
                .with_table_fallback(table_mode.into())
                .with_max_heading(max_heading)
                .with_heading_config(heading_config);

            if section_markers {
                options = options.with_section_markers(undoc::SectionMarkerStyle::Comment);
            }

            if let Some(mode) = cleanup {
                if let Some(preset) = mode.to_preset() {
                    options = options.with_cleanup_preset(preset);
                }
            }

            let markdown = undoc::render::to_markdown(&doc, &options)?;

            pb.finish_and_clear();
            write_output(output.as_ref(), &markdown)?;

            if output.is_some() {
                println!(
                    "{} Converted to Markdown: {}",
                    "✓".green().bold(),
                    output.unwrap().display()
                );
            }
        }

        Commands::Text {
            input,
            output,
            cleanup,
        } => {
            let pb = create_spinner("Parsing document...");

            let doc = undoc::parse_file(&input)?;
            pb.set_message("Rendering to text...");

            // Mirror Markdown/Convert: enable the heading analyzer for
            // consistency. The text renderer currently ignores heading levels,
            // so this is a no-op for output today, but keeps the three
            // commands aligned for future heading-aware text rendering.
            let heading_config = HeadingConfig::default().with_default_style_mapping();
            let mut options = RenderOptions::new().with_heading_config(heading_config);
            if let Some(mode) = cleanup {
                if let Some(preset) = mode.to_preset() {
                    options = options.with_cleanup_preset(preset);
                }
            }

            let text = undoc::render::to_text(&doc, &options)?;

            pb.finish_and_clear();
            write_output(output.as_ref(), &text)?;

            if output.is_some() {
                println!(
                    "{} Converted to text: {}",
                    "✓".green().bold(),
                    output.unwrap().display()
                );
            }
        }

        Commands::Json {
            input,
            output,
            compact,
        } => {
            let pb = create_spinner("Parsing document...");

            let doc = undoc::parse_file(&input)?;
            pb.set_message("Rendering to JSON...");

            let format = if compact {
                JsonFormat::Compact
            } else {
                JsonFormat::Pretty
            };
            let json = undoc::render::to_json(&doc, format)?;

            pb.finish_and_clear();
            write_output(output.as_ref(), &json)?;

            if output.is_some() {
                println!(
                    "{} Converted to JSON: {}",
                    "✓".green().bold(),
                    output.unwrap().display()
                );
            }
        }

        Commands::Info { input } => {
            let pb = create_spinner("Analyzing document...");

            let format = undoc::detect_format_from_path(&input)?;
            let doc = undoc::parse_file(&input)?;

            pb.finish_and_clear();

            println!("{}", "Document Information".cyan().bold());
            println!("{}", "─".repeat(40));
            println!(
                "{}: {}",
                "File".bold(),
                input.file_name().unwrap_or_default().to_string_lossy()
            );
            println!("{}: {:?}", "Format".bold(), format);
            println!("{}: {}", "Sections".bold(), doc.sections.len());
            println!("{}: {}", "Resources".bold(), doc.resources.len());

            if let Some(ref title) = doc.metadata.title {
                println!("{}: {}", "Title".bold(), title);
            }
            if let Some(ref author) = doc.metadata.author {
                println!("{}: {}", "Author".bold(), author);
            }
            if let Some(pages) = doc.metadata.page_count {
                println!("{}: {}", "Pages/Slides/Sheets".bold(), pages);
            }
            if let Some(ref created) = doc.metadata.created {
                println!("{}: {}", "Created".bold(), created);
            }
            if let Some(ref modified) = doc.metadata.modified {
                println!("{}: {}", "Modified".bold(), modified);
            }

            let text = doc.plain_text();
            let word_count = text.split_whitespace().count();
            let char_count = text.len();
            println!("\n{}", "Content Statistics".cyan().bold());
            println!("{}", "─".repeat(40));
            println!("{}: {}", "Words".bold(), word_count);
            println!("{}: {}", "Characters".bold(), char_count);
        }

        Commands::Extract { input, output } => {
            let pb = create_spinner("Extracting resources...");

            let doc = undoc::parse_file(&input)?;

            fs::create_dir_all(&output)?;

            let mut count = 0;
            for (id, resource) in &doc.resources {
                let raw = resource.suggested_filename(id);
                let safe_name = std::path::Path::new(&raw)
                    .file_name()
                    .unwrap_or_else(|| std::ffi::OsStr::new(&raw));
                let path = output.join(safe_name);
                fs::write(&path, &resource.data)?;
                count += 1;
            }

            pb.finish_and_clear();

            if count > 0 {
                println!(
                    "{} Extracted {} resources to {}",
                    "✓".green().bold(),
                    count,
                    output.display()
                );
            } else {
                println!("{} No resources found in document", "!".yellow().bold());
            }
        }

        Commands::Update { check, force } => {
            if let Err(e) = update::run_update(check, force) {
                eprintln!("{}: {}", "Error".red().bold(), e);
                std::process::exit(1);
            }
        }

        Commands::Version => {
            print_version();
        }
    }

    Ok(())
}

struct ConvertParams<'a> {
    input: &'a PathBuf,
    output: Option<&'a PathBuf>,
    cleanup: Option<CleanupMode>,
    formats: &'a [OutputFormat],
    no_images: bool,
    quiet: bool,
    emit_page_breaks: bool,
    include_headers_footers: bool,
    lossless: bool,
    section_markers: bool,
}

fn run_convert(p: ConvertParams<'_>) -> Result<(), Box<dyn std::error::Error>> {
    let update_rx = update::check_update_async();

    let pb = if p.quiet {
        ProgressBar::hidden()
    } else {
        create_spinner("Parsing document...")
    };

    let output_dir = match p.output {
        Some(p) => p.clone(),
        None => {
            let stem = p
                .input
                .file_stem()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();
            let parent = p.input.parent().unwrap_or(std::path::Path::new("."));
            parent.join(format!("{}_output", stem))
        }
    };

    fs::create_dir_all(&output_dir)?;

    let heading_config = HeadingConfig::default().with_default_style_mapping();
    let base = if p.lossless {
        RenderOptions::lossless()
    } else {
        RenderOptions::new()
            .with_emit_page_breaks(p.emit_page_breaks)
            .with_include_headers_footers(p.include_headers_footers)
    };
    let mut options = base
        .with_frontmatter(true)
        .with_heading_config(heading_config);
    if p.section_markers {
        options = options.with_section_markers(undoc::SectionMarkerStyle::Comment);
    }
    if let Some(mode) = p.cleanup {
        if let Some(preset) = mode.to_preset() {
            options = options.with_cleanup_preset(preset);
        }
    }

    pb.set_message("Generating output...");

    let format = undoc::detect_format_from_path(p.input)?;
    let mfw = MultiFormatWriter::new(&output_dir, p.formats, &options);

    let (summary, image_count, media_count) = match format {
        undoc::FormatType::Pptx | undoc::FormatType::Xlsx => {
            run_convert_streaming(p.input, p.no_images, &output_dir, mfw, format)?
        }
        _ => run_convert_batch(p.input, p.no_images, &output_dir, mfw)?,
    };
    let word_count = summary.word_count;

    pb.finish_and_clear();

    if !p.quiet {
        println!("{}", "Conversion Complete".green().bold());
        println!("{}", "─".repeat(40));
        println!("{}: {}", "Output".bold(), output_dir.display());
        if summary.md_path.is_some() {
            println!("  {} extract.md", "✓".green());
        }
        if summary.txt_path.is_some() {
            println!("  {} extract.txt", "✓".green());
        }
        if summary.json_path.is_some() {
            println!("  {} content.json", "✓".green());
        }
        if image_count > 0 {
            println!("  {} images/ ({} files)", "✓".green(), image_count);
        }
        if media_count > 0 {
            println!("  {} media/ ({} files)", "✓".green(), media_count);
        }

        println!("\n{}", "Statistics".cyan().bold());
        println!("{}", "─".repeat(40));
        println!("{}: {}", "Sections".bold(), summary.section_count);
        if summary.md_path.is_some() {
            println!("{}: {}", "Words".bold(), word_count);
        }
        println!(
            "{}: {} (images: {}, media: {})",
            "Resources".bold(),
            image_count + media_count,
            image_count,
            media_count
        );
    }

    if let Some(result) = update::try_get_update_result(&update_rx) {
        update::print_update_notification(&result);
    }

    Ok(())
}

fn print_version() {
    println!("{} {}", "undoc".green().bold(), env!("CARGO_PKG_VERSION"));
    println!("High-performance Microsoft Office document extraction to Markdown");
    println!();
    println!("Supported formats: DOCX, XLSX, PPTX");
    println!("Repository: https://github.com/iyulab/undoc");
}

fn create_spinner(message: &str) -> ProgressBar {
    let pb = ProgressBar::new_spinner();
    pb.set_style(
        ProgressStyle::default_spinner()
            .tick_strings(&["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"])
            .template("{spinner:.blue} {msg}")
            .unwrap(),
    );
    pb.set_message(message.to_string());
    pb.enable_steady_tick(std::time::Duration::from_millis(100));
    pb
}

type ConvertResult = (writer::WriteSummary, usize, usize);

fn run_convert_batch(
    input: &std::path::Path,
    no_images: bool,
    output_dir: &std::path::Path,
    mfw: MultiFormatWriter<'_>,
) -> Result<ConvertResult, Box<dyn std::error::Error>> {
    let doc = undoc::parse_file(input)?;
    let summary = mfw.write_document(&doc)?;
    let (image_count, media_count) = extract_resources_to_dir(&doc, no_images, output_dir)?;
    Ok((summary, image_count, media_count))
}

fn run_convert_streaming(
    input: &std::path::Path,
    no_images: bool,
    output_dir: &std::path::Path,
    mfw: MultiFormatWriter<'_>,
    doc_format: undoc::FormatType,
) -> Result<ConvertResult, Box<dyn std::error::Error>> {
    use std::ops::ControlFlow;
    use undoc::{parse_file_streaming, ParseEvent, SectionStreamOptions};

    let stream_opts = SectionStreamOptions {
        lenient: false,
        extract_resources: !no_images,
    };

    let mut sw: Option<StreamingWriter> = None;
    let mut image_count = 0usize;
    let mut media_count = 0usize;
    let images_dir = output_dir.join("images");
    let media_dir = output_dir.join("media");

    parse_file_streaming(input, stream_opts, |event| {
        match event {
            ParseEvent::DocumentStart { image_map, .. } => {
                match mfw.open_streaming(doc_format, image_map) {
                    Ok(mut writer) => {
                        let _ = writer.write_json_start();
                        sw = Some(writer);
                    }
                    Err(_) => return ControlFlow::Break(()),
                }
            }
            ParseEvent::SectionParsed(section) => {
                if let Some(ref mut writer) = sw {
                    let _ = writer.write_section(section);
                }
            }
            ParseEvent::SectionFailed { .. } | ParseEvent::DocumentEnd => {}
            ParseEvent::ResourceExtracted { name, data } => {
                let safe_name = std::path::Path::new(&name)
                    .file_name()
                    .unwrap_or_else(|| std::ffi::OsStr::new(&name));
                let ext = std::path::Path::new(&name)
                    .extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("");
                let is_image = matches!(
                    ext.to_lowercase().as_str(),
                    "png" | "jpg" | "jpeg" | "gif" | "bmp" | "tiff" | "tif" | "wmf" | "emf" | "svg"
                );
                if is_image {
                    let _ = fs::create_dir_all(&images_dir);
                    let _ = fs::write(images_dir.join(safe_name), &data);
                    image_count += 1;
                } else {
                    let _ = fs::create_dir_all(&media_dir);
                    let _ = fs::write(media_dir.join(safe_name), &data);
                    media_count += 1;
                }
            }
        }
        ControlFlow::Continue(())
    })?;

    // If streaming path didn't initialize (e.g., empty document), fall back to batch
    let summary = match sw {
        Some(writer) => writer.finish()?,
        None => {
            let doc = undoc::parse_file(input)?;
            mfw.write_document(&doc)?
        }
    };

    Ok((summary, image_count, media_count))
}

fn extract_resources_to_dir(
    doc: &undoc::Document,
    no_images: bool,
    output_dir: &std::path::Path,
) -> Result<(usize, usize), Box<dyn std::error::Error>> {
    let mut image_count = 0;
    let mut media_count = 0;

    if !no_images && !doc.resources.is_empty() {
        let images_dir = output_dir.join("images");
        let media_dir = output_dir.join("media");

        for (id, resource) in &doc.resources {
            let raw = resource.suggested_filename(id);
            let safe_name = std::path::Path::new(&raw)
                .file_name()
                .unwrap_or_else(|| std::ffi::OsStr::new(&raw));
            if resource.is_image() {
                fs::create_dir_all(&images_dir)?;
                fs::write(images_dir.join(safe_name), &resource.data)?;
                image_count += 1;
            } else {
                fs::create_dir_all(&media_dir)?;
                fs::write(media_dir.join(safe_name), &resource.data)?;
                media_count += 1;
            }
        }
    }

    Ok((image_count, media_count))
}

fn write_output(path: Option<&PathBuf>, content: &str) -> Result<(), Box<dyn std::error::Error>> {
    match path {
        Some(p) => {
            fs::write(p, content)?;
        }
        None => {
            let stdout = io::stdout();
            let mut handle = stdout.lock();
            writeln!(handle, "{}", content)?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cli_parse() {
        use clap::CommandFactory;
        Cli::command().debug_assert();
    }
}
