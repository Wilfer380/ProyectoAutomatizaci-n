# SATO WS408 machine setup and print-control guide

This guide separates what the application controls from what Windows/the printer controls. That distinction is essential when diagnosing whether a physical print issue comes from software or from the machine setup.

## Main rule

If the preview looks correct but the physical label is shifted, cut, too light, too dark, or does not print, the issue is usually in the **driver, Windows spooler, sensor, calibration, or physical media**, not in the application layout.

## What the application controls

| Control | Applied value |
|---|---|
| Target printer | `SATO WS408` |
| Logical label size | `48 mm x 23 mm` |
| Resolution | `203 DPI` |
| Sent margins | `0 mm` |
| Rendering | Same `LabelRenderer` for preview and printing |
| Batch strategy | One independent print job per label (`separate_jobs=True`) |
| Safety | Mandatory preview before printing |

## What Windows / the machine controls

| Area | Must be validated on the machine |
|---|---|
| Driver | Compatible SATO WS4 driver/Printer Utility installed. |
| Name | Windows queue must be named exactly `SATO WS408`. |
| Stock/size | Custom stock set to `48 mm x 23 mm`. |
| Sensor | Gap/label sensor calibrated for the installed roll. |
| Calibration | Printer recognizes the start/end of each label. |
| Darkness/speed | Driver or printer settings make text readable without burning. |
| Spooler | Windows queue has no stuck jobs. |
| Scaling | No “fit to page” or automatic scaling. |

## IT/support checklist

1. Install the official or compatible **SATO WS4** driver.
2. Create/verify the Windows printer queue with exact name: `SATO WS408`.
3. In printer preferences, create or select custom stock:
   - width: `48 mm`;
   - height: `23 mm`;
   - margins: `0 mm` if the driver allows it;
   - resolution: `203 DPI`.
4. Verify media type is label/gap, not continuous paper, if that applies to the roll.
5. Calibrate media/sensor from SATO Printer Utility or the printer panel/buttons.
6. Clear the Windows print queue before large tests.
7. Print a hardware/driver test from the SATO utility if available.
8. Open the app and print **1 label** first.
9. If one label is correct, test a small batch of **3 to 5 labels**.
10. If the small batch is stable, then print the full batch.

## Fast diagnosis

| Symptom | Likely source | What to do |
|---|---|---|
| App says printer is missing | Name mismatch or driver missing. | Rename queue to `SATO WS408` or reinstall driver. |
| Preview looks wrong | Data/layout/app. | Check Excel, anchored image, or `LabelRenderer`. |
| Preview is right, paper is shifted | Stock, origin, margin, or driver calibration. | Check 48x23 stock, margins, and calibration. |
| First label OK, later labels drift | Driver/spooler/page origin. | Keep `separate_jobs=True`; clear queue; review driver. |
| Output too light or too dark | Darkness, speed, ribbon/media. | Adjust darkness/speed in driver/SATO utility. |
| Nothing prints but app shows no error | Spooler, paused/offline printer, or driver. | Check Windows queue and physical state. |
| Cut occurs inside the label | Sensor/gap not calibrated or wrong stock. | Recalibrate media and confirm 48x23 mm. |

## Physical acceptance evidence

Before approving a machine for production, keep this evidence:

- [ ] Screenshot of Windows showing the `SATO WS408` queue.
- [ ] Screenshot or note of printer preferences with `48 mm x 23 mm` stock.
- [ ] Correct physical 1-label test.
- [ ] Correct physical 3–5 label batch without accumulated drift.
- [ ] Confirmation that the user knows how to preview and reject before printing.

## When to change code and when not to

| Case | Correct action |
|---|---|
| Preview and paper are both wrong | Review app/layout/data. |
| Preview is right but paper is wrong | Review machine/driver/calibration. |
| Only one machine fails | Do not change global layout; fix that machine. |
| Every machine fails with the same pattern | Review `LabelRenderer`, DPI, or logical size. |

## Operational recommendation

Keep one approved physical label as the reference sample. When a new machine is installed, compare against that sample before printing large batches.
