For **Ubuntu 24.04**, to enable PDF export via LibreOffice headless for excel-agent-tools, you need:

## Minimal Installation (Server/Headless)

```bash
sudo apt update
sudo apt install libreoffice -y
```

This installs the full suite (Writer, Calc, Impress, Draw, Base, Math) from Ubuntu's universe repository. The package includes:

- `libreoffice-core` — Base engine
- `libreoffice-calc` — Spreadsheet component (required for .xlsx → PDF)
- `libreoffice-writer` — Document component
- `libreoffice-common` — Shared resources
- `default-jre` — Java runtime (for Base database)

**Size**: Approximately 400-500 MB installed

**Version on Ubuntu 24.04**: LibreOffice 24.2.x (from universe repository)

## Verification

```bash
libreoffice --version
# Expected: LibreOffice 24.2.x.x

which soffice
# Expected: /usr/bin/soffice

# Test headless conversion
soffice --headless --convert-to pdf test.xlsx --outdir /tmp/
```

## For Newer Version (Optional)

If you need LibreOffice 25.x or newer on Ubuntu 24.04:

```bash
sudo add-apt-repository ppa:libreoffice/ppa -y
sudo apt update
sudo apt install libreoffice -y
```

## Critical Font Package (Recommended)

The research shows PDF export can have page break issues due to missing fonts. Install Microsoft core fonts:

```bash
sudo apt install ttf-mscorefonts-installer -y
```

Or for open-source alternatives:

```bash
sudo apt install fonts-liberation fonts-liberation2 -y
```

This prevents the font metrics discrepancy that causes column width calculation errors during PDF export.

## Minimal Headless-Only Installation

If you want only the components needed for conversion (no GUI dependencies):

```bash
sudo apt install libreoffice-calc libreoffice-writer --no-install-recommends -y
```

This reduces the install size but may miss some filters. The full `libreoffice` metapackage is recommended for reliability.

## For Your Test Environment

Based on your test output showing all 29 tests passing, LibreOffice is already installed and working. The single skipped test is expected behavior when LibreOffice isn't available in CI environments.

**Summary for Ubuntu 24.04**:
- **Required**: `sudo apt install libreoffice -y`
- **Recommended**: `sudo apt install ttf-mscorefonts-installer -y`
- **Verify**: `libreoffice --headless --version`
