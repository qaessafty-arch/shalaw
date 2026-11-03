# Complex Numbers: Properties of Complex Conjugate and Absolute Value

An 18-slide university presentation for first-year students.

- **University:** Salahaddin University-Erbil
- **School:** School of Science
- **Department:** Information Technology (IT)

## Deliverables

| File | What it is |
|---|---|
| `Complex_Numbers_Conjugate_and_Absolute_Value.pptx` | Editable PowerPoint deck. It has 18 slides, speaker notes on every slide, and click-to-reveal quiz answers on slide 17. |
| `Complex_Numbers_Conjugate_and_Absolute_Value.pdf` | PDF version. The quiz answers are shown, because a PDF cannot animate. |
| `SPEAKER_NOTES.md` | Speaker notes for all 18 slides. These are the same notes that are inside the .pptx. |
| `PRESENTATION_SCRIPT.md` | A natural, timed script (about 12–14 minutes) with delivery tips. |
| `MATH_ACCURACY_CHECKLIST.md` | Each formula, example and quiz answer, checked. |
| `source/` | The generator script (`build.js`). Rebuild the deck from here after you change names or the logo. |

## Before presenting

1. **Fill in your details.**
   - In PowerPoint, replace `[STUDENT NAME]`, `[LECTURER NAME]` and `[ACADEMIC YEAR]` on slides 1 and 18.
   - Or edit the `INFO` block at the top of `source/build.js` and rebuild. Rebuilding also updates the speaker notes.
2. **Add the official university logo.**
   - Save the official Salahaddin University-Erbil logo as `assets/su-logo.png`. A transparent PNG works best.
   - Rebuild (see below). The logo then replaces the dashed placeholder on slide 1.
   - You can also delete the placeholder in PowerPoint and use Insert → Pictures.
3. **Fonts.** The deck uses Montserrat for titles, Inter for body text and Cambria Math for equations.
   - Cambria Math ships with Office.
   - Install Montserrat and Inter, both free from Google Fonts, on the presenting computer. Otherwise PowerPoint substitutes similar fonts.

## Rebuilding

```bash
cd source
npm install
node build.js
```

This writes `Complex_Numbers_Conjugate_and_Absolute_Value.pptx` to the repository root.

To make the PDF, use **File → Export → PDF** in PowerPoint, or run LibreOffice:

```bash
soffice --headless --convert-to pdf Complex_Numbers_Conjugate_and_Absolute_Value.pptx
```

## Design system

- **Colors:**
  - Deep Indigo `#17152F`
  - Electric Purple `#7357FF`
  - Soft Lavender `#F1EDFF`
  - White `#FFFFFF`
  - Dark Text `#20243A`
  - Light Gray `#E8E8F0`
- **Slides:** dark slides frame the talk (cover, introduction, key identity, conclusion); the content slides are light.
- **Diagrams:** every complex-plane diagram is built from native PowerPoint shapes, drawn to scale, and fully editable.
