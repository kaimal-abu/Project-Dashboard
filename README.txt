PVUNL dashboard - electrical status update

Installation
1. Back up index.html, app.js and style.css.
2. Replace all three files together in your existing dashboard folder.
3. Keep Data-file.xlsx, Daily-updates.xlsx, Calendar-tasks.xlsx and all other files in place.
4. Open using the existing web server and hard-refresh (Ctrl+F5).

Electrical Status matches the accepted daily report:
- Three summary tiles: 11 kV drives, 3.3 kV drives, transformers.
- Drive and transformer/equipment stage bars in two panels.
- Legend is in its own wrapping row, separate from titles.
- Drive Targets table has Drive, Milestone, Target date only.
- Targets use the last entered commitment date and only the current calendar month.
- No Attention tile, timing column, schedule slip column or explanatory paragraph below bars.
- Individual equipment details remain available in expandable tables.
- Inconsistent entered counts are retained; flagged bars/details show a hover explanation.
- Electrical data read by column heading from Data-file.xlsx -> Electrical.
- Status counts, trial summaries and mixed quantities such as '3 nos' are supported.
- Blank stage data shown as missing, not as confirmed zero.
- Existing four progress charts, moving averages and all other dashboard code retained.

Layout
Manpower remains immediately before Electrical Status. Both now occupy full width,
so the electrical bars and labels have sufficient room. Two electrical panels become
one column on smaller screens. Summary tiles and detail tables also adapt to mobiles.

Tests completed
- Node JavaScript syntax check.
- HTML duplicate-ID and chart-card checks.
- Actual Electrical-sheet data tested with a DOM stub, using 19-Oct-2026.
- 25 equipment rows; three summary tiles; seven October target rows.
- November month-rollover and missing-sheet behavior tested.
- Non-electrical JavaScript verified unchanged.
Not tested in the deployed browser or on the user's server.
