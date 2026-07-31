//! Self-update functionality using GitHub releases

use colored::Colorize;
use self_update::backends::github::ReleaseList;
use self_update::cargo_crate_version;
use semver::Version;
use std::sync::mpsc;
use std::thread;
use std::time::Duration;

const REPO_OWNER: &str = "iyulab";
const REPO_NAME: &str = "undoc";
const BIN_NAME: &str = "undoc";
const CLI_CRATE_NAME: &str = "undoc-cli";

/// Platform info for asset matching
struct PlatformInfo {
    /// Human-friendly OS name (windows, linux, macos)
    os_name: &'static str,
    /// Human-friendly arch name (x86_64, aarch64)
    arch_name: &'static str,
    /// Rust target triple (x86_64-pc-windows-msvc, etc.)
    target_triple: &'static str,
    /// Archive extension (zip for Windows, tar.gz for Unix)
    archive_ext: &'static str,
}

/// Get platform info for the current system
fn get_platform_info() -> PlatformInfo {
    #[cfg(all(target_os = "windows", target_arch = "x86_64"))]
    return PlatformInfo {
        os_name: "windows",
        arch_name: "x86_64",
        target_triple: "x86_64-pc-windows-msvc",
        archive_ext: "zip",
    };

    #[cfg(all(target_os = "linux", target_arch = "x86_64"))]
    return PlatformInfo {
        os_name: "linux",
        arch_name: "x86_64",
        target_triple: "x86_64-unknown-linux-gnu",
        archive_ext: "tar.gz",
    };

    #[cfg(all(target_os = "macos", target_arch = "x86_64"))]
    return PlatformInfo {
        os_name: "macos",
        arch_name: "x86_64",
        target_triple: "x86_64-apple-darwin",
        archive_ext: "tar.gz",
    };

    #[cfg(all(target_os = "macos", target_arch = "aarch64"))]
    return PlatformInfo {
        os_name: "macos",
        arch_name: "aarch64",
        target_triple: "aarch64-apple-darwin",
        archive_ext: "tar.gz",
    };

    #[cfg(not(any(
        all(target_os = "windows", target_arch = "x86_64"),
        all(target_os = "linux", target_arch = "x86_64"),
        all(target_os = "macos", target_arch = "x86_64"),
        all(target_os = "macos", target_arch = "aarch64"),
    )))]
    {
        // Fallback for unsupported platforms
        PlatformInfo {
            os_name: std::env::consts::OS,
            arch_name: std::env::consts::ARCH,
            target_triple: "unknown",
            archive_ext: "tar.gz",
        }
    }
}

/// Generate asset name patterns to search for (in priority order)
fn get_asset_patterns(platform: &PlatformInfo, version: &str) -> Vec<String> {
    let v = version.trim_start_matches('v');
    vec![
        // Human-friendly format (preferred): undoc-windows-x86_64-v0.1.0.zip
        format!(
            "undoc-{}-{}-v{}.{}",
            platform.os_name, platform.arch_name, v, platform.archive_ext
        ),
        // Without 'v' prefix: undoc-windows-x86_64-0.1.0.zip
        format!(
            "undoc-{}-{}-{}.{}",
            platform.os_name, platform.arch_name, v, platform.archive_ext
        ),
        // Target triple format: undoc-x86_64-pc-windows-msvc-v0.1.0.zip
        format!(
            "undoc-{}-v{}.{}",
            platform.target_triple, v, platform.archive_ext
        ),
        // Target triple without 'v': undoc-x86_64-pc-windows-msvc-0.1.0.zip
        format!(
            "undoc-{}-{}.{}",
            platform.target_triple, v, platform.archive_ext
        ),
    ]
}

/// Find matching asset name from a list of asset names using exact-match fallback patterns
fn find_matching_asset(asset_names: &[String], patterns: &[String]) -> Option<String> {
    for pattern in patterns {
        if asset_names.iter().any(|name| name == pattern) {
            return Some(pattern.clone());
        }
    }
    None
}

/// Detect if installed via cargo install (binary in .cargo/bin)
fn is_cargo_install() -> bool {
    if let Ok(exe_path) = std::env::current_exe() {
        let path_str = exe_path.to_string_lossy();
        path_str.contains(".cargo") && path_str.contains("bin")
    } else {
        false
    }
}

/// Result of background update check
pub struct UpdateCheckResult {
    pub has_update: bool,
    pub latest_version: String,
    pub current_version: String,
}

/// Spawns a background thread to check for updates.
/// Returns a receiver that will contain the result when ready.
pub fn check_update_async() -> mpsc::Receiver<Option<UpdateCheckResult>> {
    let (tx, rx) = mpsc::channel();

    thread::spawn(move || {
        let result = check_latest_version();
        let _ = tx.send(result);
    });

    rx
}

/// Check for latest version without blocking (internal)
fn check_latest_version() -> Option<UpdateCheckResult> {
    let current_version = cargo_crate_version!();

    // Fetch releases from GitHub with timeout
    let releases = ReleaseList::configure()
        .repo_owner(REPO_OWNER)
        .repo_name(REPO_NAME)
        .build()
        .ok()?
        .fetch()
        .ok()?;

    if releases.is_empty() {
        return None;
    }

    let latest = &releases[0];
    let latest_version = latest.version.trim_start_matches('v');

    let current = Version::parse(current_version).ok()?;
    let latest_ver = Version::parse(latest_version).ok()?;

    Some(UpdateCheckResult {
        has_update: latest_ver > current,
        latest_version: latest_version.to_string(),
        current_version: current_version.to_string(),
    })
}

/// Try to receive update check result (non-blocking with short timeout)
pub fn try_get_update_result(
    rx: &mpsc::Receiver<Option<UpdateCheckResult>>,
) -> Option<UpdateCheckResult> {
    // Wait up to 500ms for the result
    rx.recv_timeout(Duration::from_millis(500)).ok().flatten()
}

/// Print update notification if new version available.
///
/// Writes to stderr, not stdout: `md`/`json`/`text` emit document data to
/// stdout, and a notification line there corrupts piped/redirected output
/// (e.g. invalid JSON). Diagnostics belong on stderr per Unix convention,
/// where they still surface in an interactive terminal.
pub fn print_update_notification(result: &UpdateCheckResult) {
    if result.has_update {
        eprintln!();
        eprintln!(
            "{} {} → {} available! Run '{}' to update.",
            "Update:".yellow().bold(),
            result.current_version,
            result.latest_version.green(),
            "undoc update".cyan()
        );
    }
}

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
        if is_cargo_install() {
            println!(
                "Run '{}' to update.",
                format!("cargo install {}", CLI_CRATE_NAME).cyan()
            );
        } else {
            println!("Run '{}' to update.", "undoc update".cyan());
        }
        return Ok(());
    }

    // Check installation method
    if is_cargo_install() {
        println!();
        println!(
            "{} Installed via cargo. Please run:",
            "Note:".yellow().bold()
        );
        println!(
            "  {}",
            format!("cargo install {}", CLI_CRATE_NAME).cyan().bold()
        );
        println!();
        println!(
            "{}",
            "This ensures proper integration with your Rust toolchain.".dimmed()
        );
        return Ok(());
    }

    // Perform update (GitHub Releases only)
    println!();
    println!("{}", "Downloading update...".cyan());

    // Find the correct CLI asset using exact-match patterns
    // (avoids `.contains()` collisions with FFI or other assets)
    let platform = get_platform_info();
    let patterns = get_asset_patterns(&platform, latest_version);

    let asset_names: Vec<String> = latest.assets.iter().map(|a| a.name.clone()).collect();
    let asset_name = find_matching_asset(&asset_names, &patterns);

    if asset_name.is_none() {
        println!("{}", "No matching asset found.".red());
        println!("{}", "Searched for:".dimmed());
        for p in &patterns {
            println!("  - {}", p.dimmed());
        }
        println!();
        println!(
            "{} {}",
            "Available assets:".dimmed(),
            latest
                .assets
                .iter()
                .map(|a| a.name.as_str())
                .collect::<Vec<_>>()
                .join(", ")
        );
        return Err("No compatible binary found for this platform".into());
    }

    let asset_name = asset_name.unwrap();
    println!("{} {}", "Found asset:".dimmed(), asset_name.dimmed());

    // Use direct download URL (avoids needing Accept header for API URL)
    let download_url = format!(
        "https://github.com/{}/{}/releases/download/v{}/{}",
        REPO_OWNER, REPO_NAME, latest_version, asset_name
    );

    let tmp_dir = self_update::TempDir::new()?;
    let tmp_archive_path = tmp_dir.path().join(&asset_name);
    let mut tmp_archive = std::fs::File::create(&tmp_archive_path)?;

    let mut download = self_update::Download::from_url(&download_url);
    download.show_progress(true);
    download.download_to(&mut tmp_archive)?;

    print!("Extracting archive... ");
    std::io::Write::flush(&mut std::io::stdout())?;
    let bin_name = format!("{}{}", BIN_NAME, std::env::consts::EXE_SUFFIX);
    self_update::Extract::from_source(&tmp_archive_path).extract_file(tmp_dir.path(), &bin_name)?;
    println!("Done");

    print!("Replacing binary file... ");
    std::io::Write::flush(&mut std::io::stdout())?;
    let new_exe = tmp_dir.path().join(&bin_name);
    self_update::self_replace::self_replace(new_exe)?;
    println!("Done");

    println!();
    println!(
        "{} Successfully updated to v{}!",
        "✓".green().bold(),
        latest_version
    );
    println!();
    println!("Restart undoc to use the new version.");

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn windows_x86_64() -> PlatformInfo {
        PlatformInfo {
            os_name: "windows",
            arch_name: "x86_64",
            target_triple: "x86_64-pc-windows-msvc",
            archive_ext: "zip",
        }
    }

    /// A release carries two archives whose names differ only by a `lib` prefix: the CLI
    /// executable and the shared library for the language bindings. Substring matching would
    /// let `libundoc-…` satisfy a search for `undoc-…`, and the updater would replace the
    /// executable with a library — silently, since both are valid archives. Asset names are
    /// therefore compared for equality, and this test is what keeps that true.
    #[test]
    fn cli_archive_is_chosen_over_the_library_archive() {
        let patterns = get_asset_patterns(&windows_x86_64(), "0.9.0");
        // Listed library-first, so a match that scans the release in order fails here too.
        let release = vec![
            "libundoc-windows-x86_64-v0.9.0.zip".to_string(),
            "undoc-windows-x86_64-v0.9.0.zip".to_string(),
        ];

        assert_eq!(
            find_matching_asset(&release, &patterns).as_deref(),
            Some("undoc-windows-x86_64-v0.9.0.zip")
        );
    }

    /// A release that ships only the library has nothing this updater can install. Answering
    /// "none" sends the caller down the "no matching asset" path; answering with the library
    /// would install the wrong file.
    #[test]
    fn a_library_only_release_matches_nothing() {
        let patterns = get_asset_patterns(&windows_x86_64(), "0.9.0");
        let release = vec!["libundoc-windows-x86_64-v0.9.0.zip".to_string()];

        assert_eq!(find_matching_asset(&release, &patterns), None);
    }

    /// The patterns are fallbacks in preference order, so a release using an older naming
    /// scheme still resolves — and a release offering several is answered with the preferred
    /// one rather than whichever happens to be listed first.
    #[test]
    fn naming_variants_resolve_in_preference_order() {
        let patterns = get_asset_patterns(&windows_x86_64(), "0.9.0");

        let triple_only = vec!["undoc-x86_64-pc-windows-msvc-0.9.0.zip".to_string()];
        assert_eq!(
            find_matching_asset(&triple_only, &patterns).as_deref(),
            Some("undoc-x86_64-pc-windows-msvc-0.9.0.zip")
        );

        let several = vec![
            "undoc-x86_64-pc-windows-msvc-v0.9.0.zip".to_string(),
            "undoc-windows-x86_64-v0.9.0.zip".to_string(),
        ];
        assert_eq!(
            find_matching_asset(&several, &patterns).as_deref(),
            Some("undoc-windows-x86_64-v0.9.0.zip"),
            "the human-friendly name is the preferred one"
        );
    }

    /// The tag carries a leading `v` and the asset names do not repeat it in the same place,
    /// so the prefix is stripped before the patterns are built.
    #[test]
    fn a_v_prefixed_tag_produces_the_same_patterns_as_a_bare_version() {
        let platform = windows_x86_64();
        assert_eq!(
            get_asset_patterns(&platform, "v0.9.0"),
            get_asset_patterns(&platform, "0.9.0")
        );
    }

    /// A release for a different platform must not be installed on this one.
    #[test]
    fn another_platforms_archive_is_not_matched() {
        let patterns = get_asset_patterns(&windows_x86_64(), "0.9.0");
        let release = vec![
            "undoc-linux-x86_64-v0.9.0.tar.gz".to_string(),
            "undoc-macos-aarch64-v0.9.0.tar.gz".to_string(),
        ];

        assert_eq!(find_matching_asset(&release, &patterns), None);
    }

    /// An asset from a different release is not a substitute for the one being installed.
    #[test]
    fn a_different_version_is_not_matched() {
        let patterns = get_asset_patterns(&windows_x86_64(), "0.9.0");
        let release = vec!["undoc-windows-x86_64-v0.8.0.zip".to_string()];

        assert_eq!(find_matching_asset(&release, &patterns), None);
    }
}
