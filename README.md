# Complex Numbers: Properties of Complex Conjugate and Absolute Value

A 19-slide group presentation for first-year students (cover, outline and 17 content slides). It is split between **9 presenters**: Student 1 has the cover, outline and introduction, and everyone else has 2 slides, and each presenter's name is printed at the bottom of their slides.

- **University:** Salahaddin University-Erbil
- **School:** School of Science
- **Department:** Information Technology (IT)

## Deliverables

| File | What it is |
|---|---|
| `Complex_Numbers_Conjugate_and_Absolute_Value.pptx` | Editable PowerPoint deck. It has 19 slides, speaker notes on every slide, and click-to-reveal quiz answers on slide 18. |
| `Complex_Numbers_Conjugate_and_Absolute_Value.pdf` | PDF version. The quiz answers are shown, because a PDF cannot animate. |
| `SPEAKER_NOTES.md` | Speaker notes for all 19 slides. These are the same notes that are inside the .pptx. |
| `PRESENTATION_SCRIPT.md` | The script split by presenter (Student 1 to Student 9), with the hand-over line between speakers. |
| `REFERENCES.md` | Harvard reference list, plus which source supports which slide. |
| `MATH_ACCURACY_CHECKLIST.md` | Each formula, example and quiz answer, checked. |
| `source/` | The generator script (`build.js`). Rebuild the deck from here after you change names or the logo. |

## Design styles

There are 4 versions of the deck. The content, notes, quiz animation and references are identical in all of them; only the look is different. See `styles/STYLE_PREVIEW.png` to compare them side by side.

| Style | File | Look |
|---|---|---|
| 1 Indigo | `Complex_Numbers_Conjugate_and_Absolute_Value.pptx` | White content slides with dark indigo cover, introduction, key identity and conclusion |
| 2 Light | `styles/Style_2_Light.pptx` | Every slide white, purple accents |
| 3 Dark | `styles/Style_3_Dark.pptx` | Every slide dark indigo, light text (good for dim rooms) |
| 4 Classic | `styles/Style_4_Classic.pptx` | White slides, navy accents and serif titles (Cambria, Calibri), a traditional academic look |

Each style also has a PDF next to it. Pick one and present from that file.

## Before presenting

1. **Fill in your names.**
   - The best way is to edit the `INFO` block at the top of `source/build.js` (`students`, `lecturer`, `year`) and rebuild. This fills the cover, every slide footer and the speaker notes at once.
   - You can also replace `[STUDENT 1]` to `[STUDENT 9]`, `[LECTURER NAME]` and `[ACADEMIC YEAR]` by hand in PowerPoint with Replace (Ctrl+H).
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
node build.js            # style 1 (main deck)
npm run build:all        # all 4 styles
STYLE=dark node build.js # one style: indigo, light, dark or classic
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
- **Style:** plain white content slides and a few dark slides (cover, introduction, key identity, conclusion). There are no decorative icons, shadows or stock phrases, so it reads like a real student deck.
- **Diagrams:** every complex-plane diagram is built from native PowerPoint shapes, drawn to scale, and fully editable.
