
# AZGrab

Tampermonkey userscript that auto-extracts rack positions from the weekly SharePoint workload spreadsheet. Supports **1/62AZ**, **6AZ**, **7/55AZ**, and **12AZ**.

## What It Does

- Opens the Excel file directly from SharePoint (no manual download needed)
- Filters rows by selected AZ and valid rack types (EC2, Bonsai, EBS, S3)
- Detects midday positions (above thick border separator) — included in optics, excluded from cabling
- Per-row midday toggle (🌙 Add / ☀ Remove) to manually fix incorrect entries
- Groups positions by site with 2-row gaps
- Flags lost reservations and rejected positions with visual alerts
- One-click copy as 19-column TSV — paste directly into Col A of the workload
- Separate Optics copy button for the optics workload (6-col format)

## Install

1. Install [Tampermonkey](https://www.tampermonkey.net/) browser extension
2. Click the link below — Tampermonkey will prompt you to install:

👉 **[Install AZGrab](https://raw.githubusercontent.com/morakbri/WorkloadTool//main/azgrab.user.js)**

That's it — updates are automatic.

## How to Use

1. Open the weekly workload `.xlsm` file in SharePoint
2. Click the blue **AZ** button (bottom-right corner)
3. Select an AZ from the dropdown (1/62AZ, 6AZ, 7/55AZ, 12AZ)
4. Select a sheet, then click **Grab Positions**
5. Review the results — toggle midday on/off per row if needed
6. Click **Copy Cabling (Clean)** or **Copy Optics (All)**
7. Paste into the destination workload at Col A

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 4.1.1 | 2026-10-22 | Added S3 rack type support (S3.HANUMANCV1.21 etc.); 1AZ expanded to 1/62AZ (matches 62AZ, IAD62) |
| 4.1.0 | 2026-10-22 | Renamed to AZGrab; per-row midday toggle (🌙 Add / ☀ Remove); live stats refresh on toggle; override tracking with ✎ marker |
| 4.0.3 | 2026-10-22 | Robust midday detection with 4-strategy separator scanning (MD text, thick borders, dark fills, empty rows); fixed off-by-one midday bug |
| 4.0.0 | 2026-10-22 | Multi-AZ support (1AZ, 6AZ, 7/55AZ, 12AZ) with AZ selector; dual-format matching (6AZ + IAD6 style) |
| 3.0.0 | 2026-10-22 | Midday detection switched from cell-fill colors to thick border separator; separate cabling vs. optics copy |
| 2.9.0 | 2026-09-18 | Added midday detection to exclude from cabling workload |
| 2.8.4 | 2026-09-10 | Added metronome TRUE/FALSE detection |
| 2.8.3 | 2026-09-09 | Fixed GUID extraction bug; added WOPI fetch strategy |
| 2.8.2 | 2026-09-08 | Added HTML response guard; multi-strategy auth-resilient fetch |
| 2.8.0 | 2026-09-07 | Initial multi-strategy fetch with auth headers |


