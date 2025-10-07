# Testing Checklist - BCS Compliance Monthly PPT Generator

## Quick Verification Guide

Use this checklist to verify all improvements are working correctly in generated PowerPoint files.

---

## ✅ Pre-Test Setup

- [ ] Server is running (`npm start`)
- [ ] Have test document with:
  - [ ] Multiple article categories (In the News, Federal Update, Hot Topics, Question of the Month)
  - [ ] "Go Deeper" sections in articles
  - [ ] Tables with data
  - [ ] Employer implications sections

---

## 🧪 Test Scenarios

### Test 1: Exclude Statistics Slides
**Expected:** No graph/stats slide layouts in generated PPT

- [ ] Upload test document
- [ ] Generate PPT
- [ ] Open PPT file
- [ ] **Verify:** No slides with charts, graphs, or statistics layouts
- [ ] **Pass/Fail:** ___________

---

### Test 2: Sidebar Format for Employer Implications
**Expected:** Employer implications use checklist sidebar format

- [ ] Locate employer implications slides
- [ ] **Verify:** Checklist sidebar layout used (purple background, white sidebar)
- [ ] **Verify:** Left side content area with employer implications text
- [ ] **Verify:** Right side checklist panel visible
- [ ] **Verify:** Sidebar format retained (this is the highly useful template)
- [ ] **Pass/Fail:** ___________

---

### Test 3: Category-Based Transitions
**Expected:** Transition slides vary by article category

| Article Category | Expected Transition Type | Actual | Pass/Fail |
|-----------------|-------------------------|--------|-----------|
| In the News | Standard transition | | |
| Federal Update | Alternate transition | | |
| Hot Topics | Standard transition | | |
| Question of the Month | Alternate transition | | |

- [ ] **Overall Pass/Fail:** ___________

---

### Test 4: Bulleted Outline Format
**Expected:** All content slides use bullets, not paragraphs

- [ ] Open any content slide
- [ ] **Verify:** Main points marked with bullets (•)
- [ ] **Verify:** Sub-points indented with sub-bullets (◦)
- [ ] **Verify:** NO long paragraph blocks
- [ ] **Verify:** Content optimized for presenting
- [ ] **Pass/Fail:** ___________

---

### Test 5: Table Auto-Sizing
**Expected:** Tables fit data without manual adjustment

- [ ] Locate table slides
- [ ] **Verify:** Column widths proportional to content
- [ ] **Verify:** No text overflow/cutoff
- [ ] **Verify:** Row heights accommodate data
- [ ] **Verify:** Font sizes auto-adjusted for readability
- [ ] **Pass/Fail:** ___________

---

### Test 6: "Go Deeper" Slides
**Expected:** Dedicated slides for "Go Deeper" sections

- [ ] Find slides with "Go Deeper" in title
- [ ] **Verify:** Gray section header present
- [ ] **Verify:** Content in bulleted format
- [ ] **Verify:** Nuance and context preserved from article
- [ ] **Pass/Fail:** ___________

---

### Test 7: 3-Slide Article Structure
**Expected:** Each article produces 3 slides

For each article, verify:
- [ ] **Slide 1:** Background + Applicable Employers (bulleted content)
- [ ] **Slide 2:** "Go Deeper" (if section exists in source)
- [ ] **Slide 3:** Employer Implications (checklist sidebar format)
- [ ] **Pass/Fail:** ___________

---

### Test 8: Concise Agenda Alignment
**Expected:** Agenda items match article titles

- [ ] Open agenda slide
- [ ] **Verify:** Items are concise (not verbose)
- [ ] **Verify:** Agenda items align with actual article titles
- [ ] **Verify:** Consistent naming between agenda and slides
- [ ] **Pass/Fail:** ___________

---

### Test 9: Closing Slide with Disclaimer
**Expected:** Thank You slide always present with exact disclaimer

- [ ] Navigate to last slide
- [ ] **Verify:** "THANK YOU!" title present
- [ ] **Verify:** "Q&A" subtitle present
- [ ] **Verify:** Disclaimer text exactly matches:
  > "The material presented here is for general educational purposes only and is subject to change and law, rules, and regulations. It does not provide legal or tax opinions or advice."
- [ ] **Pass/Fail:** ___________

---

## 📊 Summary Results

| Test | Pass | Fail | Notes |
|------|------|------|-------|
| 1. Exclude Statistics | [ ] | [ ] | |
| 2. Sidebar Format | [ ] | [ ] | |
| 3. Category Transitions | [ ] | [ ] | |
| 4. Bulleted Format | [ ] | [ ] | |
| 5. Table Auto-Sizing | [ ] | [ ] | |
| 6. Go Deeper Slides | [ ] | [ ] | |
| 7. 3-Slide Structure | [ ] | [ ] | |
| 8. Agenda Alignment | [ ] | [ ] | |
| 9. Closing Disclaimer | [ ] | [ ] | |

**Total Passed:** _____ / 9

**Overall Status:** ⬜ PASS | ⬜ FAIL

---

## 🐛 Issue Reporting

If any test fails, document:

1. **Test Number:** ___________
2. **What Failed:** ___________________________________________
3. **Expected Behavior:** ___________________________________________
4. **Actual Behavior:** ___________________________________________
5. **Screenshot/Evidence:** ___________________________________________

---

## 🔄 Regression Testing

After any code changes, re-run:
- [ ] All 9 core tests above
- [ ] Previous test documents to ensure no regressions
- [ ] Edge cases (empty sections, very long content, etc.)

---

**Tester Name:** ___________________
**Test Date:** ___________________
**App Version:** 2.0.1
