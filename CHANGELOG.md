# Changelog

Semantic Versioning. `MAJOR` reaches 1 when the behaviour is settled enough
to promise not to break it; 0.x is an honest statement that it may still move.

---

Clears every appearance from the active display state, where Remove Body and Component Appearances clears only what Apply Unique Colours writes.

## 0.1.5 — 2026-08-21

- Moved to its own repository. The `Source` URL in the header points at the
  new repository.
- No functional change.

## 0.1.4 — 2026-08-20

- The `Source` URL in the header now points at the renamed repository,
  `apply-colours-sw-macro`. No functional change.

## 0.1.3 — 2026-08-13

- The completion dialog reported the previous version number. See the note under
  Apply Unique Colours 0.11.2.

## 0.1.2 — 2026-08-09

- Released under the **MIT licence**, with the full text carried in the code
  itself. Header brought into line with the other macros. No functional change.

## 0.1.1

- Diagnostics off by default, now that the macro has been checked on both a part and an assembly.

## 0.1.0

- Initial version. Built on the same render material mechanism as Remove Body and Component Appearances, which had already established the two things that matter: appearance removal must be scoped through `swDisplayStateOpts_e` rather than by configuration, and `EditRebuild3` is needed before the viewport reflects the change.
- **Parts clear everything**: faces, features, bodies and the part itself. Nothing is kept, so nothing is written back and the run costs one pass.
- **Assemblies clear component appearances only**, which makes this identical to the body and component macro on an assembly. Deliberate rather than an oversight: the faces and bodies an assembly's appearances point at belong to referenced part files, and clearing those would modify documents the user never opened, including ones shared with other assemblies.
- **Asks before running.** Applying colours is reversible by the companion macro; this is not. A hand-applied face colour it clears is gone, and the macro leaves no undo record.
