//! Self-update functionality using GitHub releases

use colored::Colorize;
use indicatif::{ProgressBar, ProgressStyle};
use self_update::backends::github::{ReleaseList, Update};
use self_update::cargo_crate_version;

const REPO_OWNER: &str = "iyulab";
const REPO_NAME: &str = "undoc";
const BIN_NAME: &str = "undoc";

/// Run the update process
pub fn run_update(check_only: bool, force: bool) -> Result<(), Box<dyn std::error::Error>> {
    let current_version = cargo_crate_version!();
    println!("{} {}", "Current version:".cyan().bold(), current_version);

    println!("{}", "Checking for updates...".cyan());

    // Fetch releases from GitHub
    let releases = ReleaseList::configure()
        .repo_owner(REPO_OWNER)
        .repo_name(REPO_NAME)
        .build()?
        .fetch()?;

    if releases.is_empty() {
        println!("{}", "No releases found on GitHub.".yellow());
        return Ok(());
    }

    // Get latest release version
    let latest = &releases[0];
    let latest_version = latest.version.trim_start_matches('v');

    println!("{} {}", "Latest version:".cyan().bold(), latest_version);

    // Compare versions
    let current = semver::Version::parse(current_version)?;
    let latest_ver = semver::Version::parse(latest_version)?;

    if current >= latest_ver && !force {
        println!();
        println!("{} You are running the latest version!", "✓".green().bold());
        return Ok(());
    }

    if current < latest_ver {
        println!();
        println!(
            "{} New version available: {} → {}",
            "↑".yellow().bold(),
            current_version.yellow(),
            latest_version.green().bold()
        );
    }

    if check_only {
        println!();
        println!("Run '{}' to update.", "undoc update".cyan());
        return Ok(());
    }

    // Perform update
    println!();
    println!("{}", "Downloading update...".cyan());

    let pb = ProgressBar::new(100);
    pb.set_style(
        ProgressStyle::with_template(
            "{spinner:.green} [{bar:40.cyan/blue}] {bytes}/{total_bytes} ({eta})",
        )?
        .progress_chars("#>-"),
    );

    let target = get_target();
    let status = Update::configure()
        .repo_owner(REPO_OWNER)
        .repo_name(REPO_NAME)
        .bin_name(BIN_NAME)
        .identifier(&format!("undoc-cli-{}", target))
        .target(&target)
        .current_version(current_version)
        .show_download_progress(true)
        .no_confirm(true)
        .build()?
        .update()?;

    pb.finish_and_clear();

    match status {
        self_update::Status::UpToDate(v) => {
            println!("{} Already up to date (v{})", "✓".green().bold(), v);
        }
        self_update::Status::Updated(v) => {
            println!();
            println!("{} Successfully updated to v{}!", "✓".green().bold(), v);
            println!();
            println!("Restart undoc to use the new version.");
        }
    }

    Ok(())
}

/// Get the target triple for the current platform
fn get_target() -> String {
    #[cfg(all(target_os = "windows", target_arch = "x86_64"))]
    return "x86_64-pc-windows-msvc".to_string();

    #[cfg(all(target_os = "linux", target_arch = "x86_64"))]
    return "x86_64-unknown-linux-gnu".to_string();

    #[cfg(all(target_os = "macos", target_arch = "x86_64"))]
    return "x86_64-apple-darwin".to_string();

    #[cfg(all(target_os = "macos", target_arch = "aarch64"))]
    return "aarch64-apple-darwin".to_string();

    #[cfg(not(any(
        all(target_os = "windows", target_arch = "x86_64"),
        all(target_os = "linux", target_arch = "x86_64"),
        all(target_os = "macos", target_arch = "x86_64"),
        all(target_os = "macos", target_arch = "aarch64"),
    )))]
    {
        // Fallback: try to determine at runtime
        let arch = std::env::consts::ARCH;
        let os = std::env::consts::OS;
        match (os, arch) {
            ("windows", "x86_64") => "x86_64-pc-windows-msvc".to_string(),
            ("linux", "x86_64") => "x86_64-unknown-linux-gnu".to_string(),
            ("macos", "x86_64") => "x86_64-apple-darwin".to_string(),
            ("macos", "aarch64") => "aarch64-apple-darwin".to_string(),
            _ => format!("{}-unknown-{}", arch, os),
        }
    }
}
