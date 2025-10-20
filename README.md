# OrCAD Symbol Library Manager

A spreadsheet-style tool to batch edit OrCAD schematic symbol properties with CSV/XML round-trip, validation, and naming normalization to keep BOM-critical data clean at scale.

## Status
Active development; publishing v1.0 source and a Windows executable while newer iterations are being integrated and validated for a future release.

## Why this exists
Last-minute library property updates that feed directly into the Bill of Materials (BOM) are tedious and error-prone when done manually for each symbol.  
This tool replaces those manual edits with a fast, tabular workflow that is more reliable and efficient.

## Key Features (v6)
- Spreadsheet-like grid to view and edit symbol properties across large libraries  
- CSV/XML import and export for round-trip edits and external reviews  
- Validation helpers and quick normalization for property names and values  
- Designed to reduce late-cycle churn and BOM mismatches

## Roadmap Highlights
- Advanced validators (required fields, enums, pattern checks) and bulk transforms  
- Change review with diffs, undo/redo, and safer write-back operations  
- Preset rule packs for BOM naming standards and property mappings  
- Performance and UX improvements for handling very large libraries

## Downloads
- Source: `Lib_manager_v6.py` in this repository  
- Windows executable: packaged v6 build for quick trials without Python setup

## Quick Start
1. Export your OrCAD library to XML  
2. Open it in this tool  
3. Edit properties in the grid, validate changes  
4. Export back to XML or CSV  
5. Re-import the updated XML to OrCAD following your standard library update process

## Contribution
Contributions are welcome via issues and pull requests!  
Start with small validators, UI refinements, or documentation examples. Larger designs are better discussed in issues first.

## License and Commercial Use
This project is open for community contributions and non-commercial use.  
Commercial use or redistribution requires explicit written permission from the owner.

## Notes
v1.0 reflects the current public baseline. Newer code will be released after further stabilization, cross-platform packaging, and regression testing.

---

If you need assistance with installation or want to request a feature, please open an issue or discussion.

Thank you for your interest and contributions!
