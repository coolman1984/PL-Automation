# Offline package cache

This delivery already contains the locked Windows x64 binary packages and a
`wheelhouse-manifest.json` with SHA-256 hashes. Run
`PREPARE_OFFLINE_PACKAGES.bat` only when intentionally rebuilding the package
set on an internet-connected computer. The cache supports 64-bit Python 3.12,
3.13, and 3.14 on Windows 10/11.

Production users should run the packaged executable. This cache exists only for
deterministic offline building and private source-mode repair; it must never
install packages globally or change the system Python installation.
