#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
venv_dir="${repo_root}/.omx/.venv"

cd "${repo_root}"

echo "[lane4] building native library"
cargo build --release --features ffi

if [[ ! -x "${venv_dir}/bin/python" ]]; then
  echo "[lane4] creating local python venv at ${venv_dir}"
  python3 -m venv "${venv_dir}"
fi

echo "[lane4] ensuring pytest is available"
"${venv_dir}/bin/pip" install pytest >/dev/null

echo "[lane4] running python native-backed regressions"
"${venv_dir}/bin/python" -m pytest bindings/python/tests/test_undoc.py -vv

echo "[lane4] running csharp native-backed regressions"
dotnet test bindings/csharp/Undoc.Tests/Undoc.Tests.csproj -c Release
