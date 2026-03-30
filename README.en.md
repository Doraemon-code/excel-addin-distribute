<div align="center">

# Excel Addin Installer

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()
[![Python](https://img.shields.io/badge/python-3.8+-blue.svg)]()

A Windows desktop tool for installing and updating Excel add-ins (.xlam files) with WebDAV remote deployment support

English | **[简体中文](README.md)**

</div>

---

## Features

- **Remote Installation** — Download and install via WebDAV
- **Local Installation** — Install from local .xlam files
- **Auto Registry Config** — Automatic Excel add-in registry setup
- **VBA Version Control** — Git hooks export VBA code for versioning

---

## Workflow

### Administrator (Publisher)

#### 1. Build xlam with Ribbon UI

Create Excel add-in using [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor):

1. Open Excel → New blank workbook
2. `Alt + F11` → Insert Module → Paste VBA code
3. Save as `MyAddin.xlam` (Excel Add-in format)
4. Open with **Office RibbonX Editor**
5. Insert → Office 2010+ Custom UI Part
6. Paste CustomUI XML → Validate → Save

> **Important:** Ribbon callbacks must include `IRibbonControl` parameter:
> ```vba
> Public Sub MyCallback(control As IRibbonControl)
>     ' code here
> End Sub
> ```

> **Editing xlam correctly:**
> - ❌ Don't double-click xlam (loads as add-in, Ctrl+S won't save code)
> - ✅ Close all Excel → Open blank workbook → File → Open → Select xlam → Edit → Save

#### 2. Deploy to WebDAV

Upload to WebDAV server:

```
Addin/
├── MyAddin.xlam
└── version.json
```

**version.json:**

```json
{
  "version": "1.0.0",
  "releaseDate": "2026-03-27",
  "filename": "MyAddin.xlam",
  "changelog": "Release notes"
}
```

#### 3. Distribute Installer

Provide `ExcelAddinInstaller.exe` to users.

---

### User (Installer)

1. Run `ExcelAddinInstaller.exe`
2. Choose installation method:
   - **Remote**: Enter WebDAV URL/credentials (or use defaults)
   - **Local**: Select local .xlam file
3. (Remote) Click "Check Update" → Confirm new version → Click "Install"
4. **Restart Excel completely** to see Ribbon tab

---

## WebDAV Notes

| Item | Description |
|------|-------------|
| Version must increment | Client won't detect updates otherwise |
| Filename must match | `version.json` filename field must equal actual file name |
| Folder name consistent | Default folder is `Addin/` (configurable via `webdav_default_folder`) |
| Resources accessible | Ensure WebDAV resources are downloadable (URL, auth, permissions, certificates) |

---

## Configuration

| File | Location | Priority |
|------|----------|----------|
| `config.yaml` | Embedded in exe | Default values |
| `config.json` | `%APPDATA%\ExcelAddinInstaller\` | User overrides |

---

## Development

```bash
# Create virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run
python main.py
```

### Build

```bash
pyinstaller --onefile --windowed \
  --name ExcelAddinInstaller \
  --icon app.ico \
  --add-data "config.yaml;." \
  main.py
```

### Git Hooks (VBA Version Control)

```bash
# Windows
scripts\install_hook.bat

# Unix/Linux/macOS
sh scripts/install_hook.sh
```

Pre-commit hook automatically exports VBA code from `.xlam` files to `src/vba/` for version control.

---

## License

[MIT License](LICENSE)