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
    /// Convert a document (default command - extracts all formats)
    Convert {
        /// Input file path
        input: PathBuf,

        /// Output directory (default: <filename>_output)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Apply text cleanup
        #[arg(long)]
        cleanup: Option<CleanupMode>,
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
    // Handle default command (undoc <file> [output])
    if cli.command.is_none() {
        if let Some(input) = cli.input {
            return run_convert(&input, cli.output.as_ref(), cli.cleanup);
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
        } => {
            run_convert(&input, output.as_ref(), cleanup)?;
        }

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

/// Run the default convert command - extracts all formats to output directory
fn run_convert(
    input: &PathBuf,
    output: Option<&PathBuf>,
    cleanup: Option<CleanupMode>,
) -> Result<(), Box<dyn std::error::Error>> {
    let pb = create_spinner("Parsing document...");

    // Determine output directory
    let output_dir = match output {
        Some(p) => p.clone(),
        None => {
            let stem = input
                .file_stem()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();
            let parent = input.parent().unwrap_or(std::path::Path::new("."));
            parent.join(format!("{}_output", stem))
        }
    };

    // Create output directory
    fs::create_dir_all(&output_dir)?;

    // Parse document
    let doc = undoc::parse_file(input)?;

    // Prepare render options
    let mut options = RenderOptions::new().with_frontmatter(true);
    if let Some(mode) = cleanup {
        options = options.with_cleanup_preset(mode.into());
    }

    // Generate Markdown
    pb.set_message("Generating Markdown...");
    let markdown = undoc::render::to_markdown(&doc, &options)?;
    let md_path = output_dir.join("extract.md");
    fs::write(&md_path, &markdown)?;

    // Generate plain text
    pb.set_message("Generating text...");
    let text = undoc::render::to_text(&doc, &options)?;
    let txt_path = output_dir.join("extract.txt");
    fs::write(&txt_path, &text)?;

    // Generate JSON
    pb.set_message("Generating JSON...");
    let json = undoc::render::to_json(&doc, JsonFormat::Pretty)?;
    let json_path = output_dir.join("content.json");
    fs::write(&json_path, &json)?;

    // Extract resources
    let mut resource_count = 0;
    if !doc.resources.is_empty() {
        pb.set_message("Extracting resources...");
        let media_dir = output_dir.join("media");
        fs::create_dir_all(&media_dir)?;

        for (id, resource) in &doc.resources {
            let filename = resource.suggested_filename(id);
            let path = media_dir.join(&filename);
            fs::write(&path, &resource.data)?;
            resource_count += 1;
        }
    }

    pb.finish_and_clear();

    // Print summary
    println!("{}", "Conversion Complete".green().bold());
    println!("{}", "─".repeat(40));
    println!("{}: {}", "Output".bold(), output_dir.display());
    println!("  {} extract.md", "✓".green());
    println!("  {} extract.txt", "✓".green());
    println!("  {} content.json", "✓".green());
    if resource_count > 0 {
        println!("  {} media/ ({} files)", "✓".green(), resource_count);
    }

    // Print statistics
    let word_count = text.split_whitespace().count();
    println!("\n{}", "Statistics".cyan().bold());
    println!("{}", "─".repeat(40));
    println!("{}: {}", "Sections".bold(), doc.sections.len());
    println!("{}: {}", "Words".bold(), word_count);
    println!("{}: {}", "Resources".bold(), resource_count);

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
