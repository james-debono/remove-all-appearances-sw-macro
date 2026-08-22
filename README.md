# Remove All Appearances

A SOLIDWORKS macro that strips every appearance from the active display state,
including ones applied by hand.

Works with SOLIDWORKS 2022, 2024 and 2025.

## What it does

The blunt instrument. In a **part** it clears face, feature, body and part
appearances — everything, however it got there. In an **assembly** it clears
component appearances only, and never modifies the referenced part files.

Only the **active display state** is touched. Other display states keep their
appearances.

**It asks for confirmation before doing anything**, because it destroys work that
cannot be recreated and leaves no undo record. That prompt is deliberate: this
macro removes colours you applied deliberately by hand, not just ones a macro
wrote.

If you only want to clear the body and component level — undoing
[Apply Unique Colours](https://github.com/james-debono/apply-unique-colours-sw-macro)
while keeping hand-applied face colours — use
[Remove Body and Component Appearances](https://github.com/james-debono/remove-body-and-component-appearances-sw-macro)
instead.

## Install

**The macro on its own:** download `Remove-All-Appearances.swp` from the
[latest release](../../releases/latest), then run it with **Tools > Macro > Run**,
or add it to a toolbar with **Tools > Customize > Commands > Macro**.

**With [MacroDeck](https://github.com/james-debono/macrodeck-sw-addin):** get the
[MacroDeck Collection](https://github.com/james-debono/macrodeck-collection-sw-macro-library/releases/latest),
which packages this macro with its icon and hover text alongside every other macro
in the set. Point MacroDeck at the unzipped folder and it appears as a button.

## Using it

Open a part or assembly and run the macro. Confirm the prompt. It reports how many
appearances it cleared when it finishes.

## Limitations

- **No undo.** Appearance changes leave no undo record, which is why the macro
  asks first.
- Only the active display state is affected.
- In an assembly, only component appearances are cleared. Appearances living
  inside the referenced part files are untouched — open the part and run it there.

## Related macros

- [Apply Unique Colours](https://github.com/james-debono/apply-unique-colours-sw-macro)
  — colours every geometrically unique body
- [Remove Body and Component Appearances](https://github.com/james-debono/remove-body-and-component-appearances-sw-macro)
  — the targeted removal, which keeps face and feature colours

## Building from source

`src\Remove-All-Appearances.vba` is the readable source. A `.swp` is a binary VBA
project, so it can only be produced from inside SOLIDWORKS — there is no build
step. Open the `.swp` via **Tools > Macro > Edit**, paste the source in, and save.

Technical detail is in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Licence

MIT — see [LICENSE](LICENSE). Free to use, modify and share. The full licence text
is also carried inside the macro itself, so a `.swp` passed on by itself still
carries its licence.

Written by James Debono, with AI assistance. Everything here was tested by
hand in SOLIDWORKS — nothing that touches the API can be verified any other way.

## Trademarks

SOLIDWORKS is a registered trademark of Dassault Systèmes SolidWorks Corporation.
This project is independent: it is not affiliated with, endorsed by, or sponsored
by Dassault Systèmes, and uses only the published SOLIDWORKS API.
