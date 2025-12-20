//! undoc CLI - Microsoft Office document extraction tool
//!
//! A command-line tool for extracting content from DOCX, XLSX, and PPTX files.

mod update;

use clap::{Parser, Subcommand, ValueEnum};
use colored::*;
use indicatif::{ProgressBar, ProgressStyle};
use std::fs;
use std::io::{self, Write};
use std::path::PathBuf;
use undoc::render::{CleanupPreset, JsonFormat, RenderOptions, TableFallback};

/// Microsoft Office document extraction to Markdown, text, and JSON
#[derive(Parser)]
#[command(
    name = "undoc",
    author = "iyulab",
    version,
    about = "Extract content from Office documents",
    long_about = "undoc - High-performance Microsoft Office document extraction tool.\n\n\
                  Converts DOCX, XLSX, and PPTX files to Markdown, plain text, or JSON."
)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
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

        /// Maximum heading level (1-6)
        #[arg(long, default_value = "6")]
        max_heading: u8,
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
    /// Minimal cleanup
    Minimal,
    /// Standard cleanup (default)
    Standard,
    /// Aggressive cleanup
    Aggressive,
}

impl From<CleanupMode> for CleanupPreset {
    fn from(mode: CleanupMode) -> Self {
        match mode {
            CleanupMode::Minimal => CleanupPreset::Minimal,
            CleanupMode::Standard => CleanupPreset::Default,
            CleanupMode::Aggressive => CleanupPreset::Aggressive,
        }
    }
}

fn main() {
    let cli = Cli::parse();

    if let Err(e) = run(cli) {
        eprintln!("{}: {}", "Error".red().bold(), e);
        std::process::exit(1);
    }
}

fn run(cli: Cli) -> Result<(), Box<dyn std::error::Error>> {
    match cli.command {
        Commands::Markdown {
            input,
            output,
            frontmatter,
            table_mode,
            cleanup,
            max_heading,
        } => {
            let pb = create_spinner("Parsing document...");

            let doc = undoc::parse_file(&input)?;
            pb.set_message("Rendering to Markdown...");

            let mut options = RenderOptions::new()
                .with_frontmatter(frontmatter)
                .with_table_fallback(table_mode.into())
                .with_max_heading(max_heading);

            if let Some(mode) = cleanup {
                options = options.with_cleanup_preset(mode.into());
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

            let mut options = RenderOptions::new();
            if let Some(mode) = cleanup {
                options = options.with_cleanup_preset(mode.into());
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
                let filename = resource.suggested_filename(id);
                let path = output.join(&filename);
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
                println!(
                    "{} No resources found in document",
                    "!".yellow().bold()
                );
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
