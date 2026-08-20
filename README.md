# UIO IFAC 2025

Open only these files:

- [Paper (PDF)](ifac/UIO_IFAC.pdf)
- [Presentation (PDF)](ifac/UIO_IFAC_15min_beamer.pdf)
- [Presenter view: slides + speaker notes (PDF)](ifac/UIO_IFAC_15min_presenter_notes.pdf)
- [Internal Layer-2 Q&A (PDF)](ifac/INTERNAL_QA_LAYER2.pdf)

Sources, backups, figures, and LaTeX build files are not needed for reading.

## Presenter mode (slides + notes)

1. Set Windows display mode to **Extend**.
2. Install Pympress once:

   ```powershell
   winget install --id Cimbali.pympress --exact --source winget
   ```

3. From this folder, launch the presenter-notes PDF:

   ```powershell
   pympress --notes=right ".\ifac\UIO_IFAC_15min_presenter_notes.pdf"
   ```

Keep the presenter window on the laptop; Pympress sends the slide-only view to the projector. Press `S` if the two screens are reversed.
