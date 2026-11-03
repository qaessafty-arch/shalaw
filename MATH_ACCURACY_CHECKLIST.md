# Mathematical Accuracy Checklist

Every formula, example and quiz answer in the deck was checked by hand and again with a short Python script.

## Definitions

| Item | Slide | Statement | Status |
|---|---|---|---|
| Imaginary unit | 2 | i² = −1 | OK |
| Standard form | 3 | z = a + bi, with a = Re(z) and b = Im(z) (b is real and does not include i) | OK |
| Conjugate | 6 | z = a + bi ⇒ z̄ = a − bi | OK |
| Absolute value | 10 | \|z\| = √(a² + b²), the distance from the origin (Pythagoras) | OK |

## Examples

| Slide | Calculation | Result | Status |
|---|---|---|---|
| 2 | 2² = 4, (−2)² = 4, 0² = 0 | real squares are never negative | OK |
| 3 | 3 + 4i → Re 3, Im 4 · 2 − 5i → Re 2, Im −5 | - | OK |
| 4 | 5 + 2i → Re 5, Im 2 · −3 + 7i → Re −3, Im 7 | - | OK |
| 5 | 3 + 4i plotted at (3, 4) | - | OK |
| 6–7 | conj(3 + 4i) = 3 − 4i, the reflection of (3, 4) to (3, −4) | - | OK |
| 8 (rule 1) | z = 2 + 3i → z̄ = 2 − 3i → conjugate again = 2 + 3i | z̿ = z | OK |
| 8 (rule 2) | z₁ = 1 + 2i, z₂ = 3 + i: z₁ + z₂ = 4 + 3i → 4 − 3i; (1 − 2i) + (3 − i) = 4 − 3i | equal | OK |
| 8 (rule 3) | z₁ = 1 + i, z₂ = 1 + 2i: z₁z₂ = 1 + 2i + i + 2i² = −1 + 3i → −1 − 3i; (1 − i)(1 − 2i) = 1 − 2i − i + 2i² = −1 − 3i | equal | OK |
| 9 | conj(6 + 2i) = 6 − 2i · conj(−4 − 3i) = −4 + 3i · conj(7) = 7 | - | OK |
| 10 | \|−5\| = 5 | - | OK |
| 11–12 | \|3 + 4i\| = √(9 + 16) = √25 = 5 | 5 | OK |
| 12 | \|5 + 12i\| = √(25 + 144) = √169 = 13 | 13 | OK |
| 13 | \|3 + 4i\| = 5 · \|0\| = 0 · \|3 − 4i\| = 5 · \|(3 + 4i)·2\| = \|6 + 8i\| = 10 = 5 × 2 | - | OK |
| 14 | (3 + 4i)(3 − 4i) = 9 − 12i + 12i − 16i² = 9 + 16 = 25 = 5² | z z̄ = \|z\|² | OK |
| 14 | 1 ÷ (3 + 4i) = (3 − 4i) ÷ 25 | - | OK |
| 15 | z = 3 + 4i: z̄ = 3 − 4i, \|z\| = 5 | - | OK |

## Quiz answers (slide 17)

| Q | Question | Answer | Check | Status |
|---|---|---|---|---|
| 1 | conjugate of 4 + 6i | 4 − 6i | flip the sign of 6i | OK |
| 2 | \|3 + 4i\| | 5 | √(9 + 16) = √25 | OK |
| 3 | conjugate of −2 − 5i | −2 + 5i | real part unchanged | OK |
| 4 | \|5 + 12i\| | 13 | √(25 + 144) = √169 | OK |

## Properties stated (no proofs, as required)

- z̿ = z  OK
- conj(z₁ + z₂) = z̄₁ + z̄₂  OK
- conj(z₁z₂) = z̄₁ z̄₂  OK
- \|z\| ≥ 0  OK
- \|z\| = 0 ⇔ z = 0  OK
- \|z̄\| = \|z\|  OK
- \|z₁z₂\| = \|z₁\|\|z₂\|  OK
- z z̄ = \|z\|²  OK

## Diagrams

- **Slide 5:** point (3, 4), with guides drawn to both axes.  OK
- **Slide 7:** z at (3, 4) and z̄ at (3, −4), the same distance from the real axis.  OK
- **Slide 10:** the triangle labels a, b, c, with c the hypotenuse.  OK
- **Slide 11:** the triangle has legs 3 (horizontal) and 4 (vertical) and a right angle at (3, 0); the hypotenuse is 5.  OK
- **Slide 16:** the arrow length is labelled magnitude and the angle φ is labelled phase. The dashed reference wave and the solid wave differ only by a phase shift.  OK

The graphs are drawn to scale: one grid unit is one unit on both axes.

## Content requirements

- Exactly **18 slides**.  OK
- English only, with no Kurdish or Arabic text.  OK
- Beginner level: no proofs and no advanced formulas.  OK
- References are real and in Harvard style (slide 18 and REFERENCES.md), with an in-text citation on each content slide.  OK
- University details: Salahaddin University-Erbil, School of Science, IT Department.  OK
- 9 presenters, 2 slides each, with names in the slide footers.  OK
- **Official logo:** the sandbox could not download it, so slide 1 has a marked placeholder. See README.md for how to add the official file.
