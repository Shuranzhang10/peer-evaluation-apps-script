# Peer Evaluation for Group Projects (Google Apps Script)

A lightweight peer-evaluation web app for course group projects, built with Google Apps Script and a Google Sheet.
No server and no database are needed, and the whole thing lives in your own Google account.
It was first built for **AMS 325 (Stony Brook University) final project presentations**, and the scoring display is modeled on the VIP program's peer evaluation.

## What it does

Students open one link, pick their group and name, and verify with their student ID.

**Part 1 – Teammate evaluation (confidential)**
- Each student rates every teammate on several 1–5 questions (default: contribution, reliability, communication).
- They write one sentence on what each teammate worked on, and one on their own main contribution.
- They can add an optional confidential comment.
- Only the instructor/TA (the Sheet owner) can see these answers.

**Part 2 – Presentation evaluation**
- Students rate each presenting group (never their own) on up to 5 questions (1–5).
- They can add an optional anonymous comment.
- Each group is submitted separately, so students can rate right after each talk.

**Results (in the Google Sheet, one click)**
- Each cell shows *score / that rater's own average*.
- Cells are colored by how far the score is from the rater's usual rating (z-score), so "always gives 5" and "always gives 2" habits are accounted for:
  - dark green: more than +1 SD
  - light green: +0.5 to +1 SD
  - white: within ±0.5 SD
  - yellow: −0.5 to −1 SD
  - red: below −1 SD
- **Team Results** shows, for each group, every member's average, the self-reported contribution, what teammates say they did, who did not submit, and VIP-style rater × ratee matrices.
- **Presentation Results** shows each group's average per question (color = rater-adjusted), an adjusted rank, and all anonymous comments.
- **Presentation Detail** shows every rater × group score.

## Files

| File | What it is |
|---|---|
| `Code.gs` | Server code: roster, questions, saving responses, and building the results. |
| `Index.html` | The page students see (mobile friendly). |

## Setup (about 10 minutes)

1. Create a new Google Sheet, then open **Extensions → Apps Script**.
2. Replace the default code with `Code.gs`.
3. Add an HTML file named `Index` (**+ → HTML**) and paste `Index.html` into it. Save.
4. Edit the **`ROSTER`** in `Code.gs` with your students (name, ID, group number). Keep the group-0 test students.
5. Select the `setup` function and click **Run**, then authorize it. This creates the `Settings`, `Teammate`, `Self` and `Presentation` sheets.
6. Click **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone within <your university>** (recommended), or **Anyone**
7. Share the web-app URL with students (Brightspace, or a QR code on the projector).

After any code change, click **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. The URL stays the same.

## Day-to-day settings (no redeploy needed)

Edit the **Settings** sheet:

| Setting | Example | Meaning |
|---|---|---|
| `TEAMMATE_OPEN` | `TRUE` | Turn Part 1 on or off |
| `PRESENTATION_OPEN` | `TRUE` | Turn Part 2 on or off |
| `PRESENTATION_GROUPS` | `1-7`, `8-13`, `1,2,5` | Which groups can be rated right now |

> Tip: keep the `PRESENTATION_GROUPS` cell formatted as **Plain text**, because Sheets may turn `1-7` into a date. The code also handles this case.

## Changing the questions

All of the wording lives at the top of `Code.gs`:

- `TEAM_RATING_QUESTIONS`: the teammate questions (1–5 scale, with `low` / `high` labels).
- `TEAM_TEXT_QUESTION` and `SELF_QUESTION`: the open-ended teammate questions.
- `PRES_QUESTIONS`: the presentation questions (up to 5 recommended).

Edit them, save, and deploy a new version.

## Testing (for instructors)

Open the web-app URL with **`?test=1`** at the end. A **TEST (instructors)** group then appears, with these test students:

| Test student | ID |
|---|---|
| Test Student A | `000000001` |
| Test Student B | `000000002` |
| Test Student C | `000000003` |

Students never see this group, and test responses are excluded from all results.
To see the colored results with test data, use **Peer Eval → Preview results incl. test data** (run **Build results** again afterwards for the real results).
To remove test responses, use **Peer Eval → Clear test data** in the Sheet.

## Viewing results

In the Sheet, use **Peer Eval → Build results**. You can re-run it at any time; the result sheets are rebuilt from the raw data.

## Privacy notes

- Student IDs stay on the server side and are never sent to the browser.
- Do **not** commit a real class roster to a public repository. Keep the real `ROSTER` only in your own Apps Script project.
