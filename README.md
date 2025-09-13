# Excel VBA – Highlight & Count Toolkit

Two handy Excel VBA macros to quickly **highlight and count** staff entries (e.g., by Name or Department) and to **summarize counts by group**. The result is written into the **first empty column to the right** of your existing data, keeping your source table intact.

> Note: Use on anonymized/sample data in public repos.

---

## Features

- **Highlight & Count Matches**
  - Choose a field by its **header name** (e.g., `Department`) or **column letter** (e.g., `C`)
  - Search by **Exact** or **Contains** matches
  - Highlights matching rows and shows a **count + parameters** in the next empty column

- **Summarize Counts by Field**
  - Produces a quick **frequency table** (e.g., how many entries per Department)
  - Output placed in the **next empty column** to the right

- **Clear Highlights**
  - Resets previous row highlights and bolding in the UsedRange

---

## Installation

1. Open your workbook
2. Press `ALT + F11` (VBA Editor)
3. *Insert* → *Module*
4. Paste the code from `Module1.bas` (or copy this repository’s code)
5. Save the workbook as **.xlsm** to keep macros

---

## Data Assumptions

- First row of the *UsedRange* contains headers (e.g., `Name`, `Department`, `EmployeeID`, etc.)
- Data begins on the next row

---

## How to Use

### 1) Highlight & Count
- `ALT + F8` → run **HighlightAndCountMatches**
- When prompted:
  - Enter a **field** (header or column letter)
  - Choose **1** for **Exact** or **2** for **Contains**
  - Enter search term(s) (comma-separated for multiple)
- The macro highlights matching rows and writes:
  - `Match Count`
  - `Field`, `Mode`, `Terms`
  - …to the **first empty column** to the right of your data

### 2) Summarize Counts by Field
- `ALT + F8` → run **SummarizeCountsByField**
- Enter the field (header or letter) to group by (e.g., `Department`)
- A two-column table (`Value`, `Count`) appears in the **next empty column** area

### 3) Clear Highlights (Optional)
- `ALT + F8` → run **ClearHighlights** to reset highlight/bold formatting

---

## Examples

- Highlight all rows where **Department contains** “Operations”
- Count all entries where **Name equals** “Jane Doe”
- Summarize how many staff are in each **Department**

---

## Notes

- Matching is **case-insensitive**
- The “first empty column to the right” is computed from the current UsedRange
- You can re-run the tools multiple times; summaries append in new “next” columns

---

## Author
**Sander Soares**  
BSc (Hons) in Computing IT | Networking & Cloud Enthusiast  
---
