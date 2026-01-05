# VBA Source Workflow Manual

This project uses an **external VBA source workflow** to allow clean editing in VS Code, version control (Git), and reproducible Excel builds.

Excel itself is treated as a **runtime container**, not the primary source editor.

---

## Folder Structure

```
tools/
└─ src/
   ├─ *.bas / *.cls / *.frm        ← VBA source files (source of truth)
   └─ _internal/
      ├─ ThisWorkbook.cls
      ├─ Sheet_*.cls
      └─ notes / docs
```

### `tools/src/`
- Contains **all importable VBA code**
- These files are:
  - edited in VS Code
  - versioned in Git
  - imported into Excel via macro
- Supported file types:
  - `.bas` (standard modules)
  - `.cls` (class modules)
  - `.frm` (+ `.frx`) (UserForms)

### `tools/src/_internal/`
- Contains **document-level VBA code**:
  - `ThisWorkbook`
  - worksheet modules (`Sheet1`, `Sheet_Pokedex`, …)
- These files are **for reference and documentation only**
- ❗ They are **NOT imported back into Excel**
  - Excel requires document modules to live directly inside the workbook

---

## Importing VBA into Excel

To load all VBA source files from `tools/src/` into the current workbook:

```vba
ImportAllVba
```

What it does:
- Imports all `.bas`, `.cls`, `.frm` files from `tools/src/`
- Overwrites existing modules with the same name
- Ignores:
  - subfolders
  - `_internal`
  - non-VBA files (e.g. `.txt`, `.md`)

---

## Exporting VBA from Excel

To export all VBA from the workbook back to disk:

```vba
ExportAllVba
```

What it does:
- Asks for confirmation before running
- Exports:
  - standard modules, classes, forms → `tools/src/`
  - document modules (`ThisWorkbook`, sheets) → `tools/src/_internal/`
- Existing files are overwritten

> ⚠️ Export is intended mainly as a **sync / safety tool**.  
> The primary source of truth should always be `tools/src/`.

---

## Required Excel Security Setting

Both import and export macros require Excel to allow programmatic access to the VBA project.

### How to enable it:
1. Open **Excel**
2. Go to **File → Options**
3. Select **Trust Center**
4. Click **Trust Center Settings**
5. Open **Macro Settings**
6. Enable:
   - ✅ **Trust access to the VBA project object model**

Without this setting:
- Import/export macros will fail
- Excel will block access to `VBProject`

---

## Recommended Usage Rules

- ✅ Edit VBA only in `tools/src/`
- ❌ Do not manually edit VBA inside Excel (except document event stubs)
- ✅ Treat Excel files as build/runtime artifacts
- ✅ Use ImportAllVba during development
- ✅ Use ExportAllVba only when synchronizing or auditing

---

## Summary

- `tools/src` = **source of truth**
- Excel = **runtime container**
- `_internal` = **document-level reference only**
- Import/Export macros keep everything in sync

This setup enables clean VBA development with modern tooling while respecting Excel’s architectural limitations.
