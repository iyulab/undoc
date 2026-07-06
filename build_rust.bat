@echo off
set CARGO_TARGET_DIR=F:\Elements_only\ephem\apps\rust\auto_generated\undoc_workspace

echo ====================================================
echo BUILDING UNDOC Crate
echo ====================================================

cargo build

:: for production release 
:: cargo build --release