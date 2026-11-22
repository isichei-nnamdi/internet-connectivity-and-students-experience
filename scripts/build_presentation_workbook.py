"""Build the presentation workbook (reports/NSUK_Connectivity_Presentation.xlsx).

Every summary number is a formula over the 'Clean Data' tab, so the workbook stays
traceable after it is converted to Google Sheets. Statistical test results, which
spreadsheets cannot reproduce, are imported from the notebook outputs with a source note.

Run after the notebook:  python scripts/build_presentation_workbook.py
"""
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
df = pd.read_csv(ROOT / "data/processed/survey_clean.csv")
TAB = ROOT / "reports/tables"
OUT = ROOT / "reports/NSUK_Connectivity_Presentation.xlsx"

# ---------------------------------------------------------------- style
FONT = "Arial"
NAVY, BLUE, ORANGE, AQUA = "184F95", "2A78D6", "EB6834", "1BAF7A"
INK, INK2, BAND, LINE = "0B0B0B", "52514E", "F3F6FA", "D9DDE3"
F_TITLE = Font(name=FONT, size=16, bold=True, color=INK)
F_SUB = Font(name=FONT, size=10, italic=True, color=INK2)
F_H2 = Font(name=FONT, size=12, bold=True, color=NAVY)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=FONT, size=10, color=INK)
F_BOLD = Font(name=FONT, size=10, bold=True, color=INK)
F_KPI = Font(name=FONT, size=14, bold=True, color=NAVY)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_BAND = PatternFill("solid", fgColor=BAND)
FILL_NOTE = PatternFill("solid", fgColor="FFF4E5")
THIN = Border(bottom=Side(style="thin", color=LINE))
WRAP = Alignment(wrap_text=True, vertical="top")
PCT = "0%"

wb = Workbook()
wb.remove(wb.active)

def sheet(name, title, subtitle, widths):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A2"] = title, subtitle
    ws["A1"].font, ws["A2"].font = F_TITLE, F_SUB
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws

def header(ws, row, labels, col=1):
    for j, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=lab)
        c.font, c.fill = F_HEAD, FILL_HEAD
        c.alignment = Alignment(horizontal="center" if j else "left", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30

def body(ws, r0, r1, c0, c1, fmts=None, band=True):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            if not (cell.font and cell.font.bold): cell.font = F_BODY
            cell.border = THIN
            if band and (r - r0) % 2 == 1:
                cell.fill = FILL_BAND
            if fmts and c in fmts:
                cell.number_format = fmts[c]
            if c > c0:
                cell.alignment = Alignment(horizontal="center", vertical="center")

def h2(ws, ref, text):
    ws[ref] = text; ws[ref].font = F_H2

def note_box(ws, ref_range, text):
    ws.merge_cells(ref_range)
    c = ws[ref_range.split(":")[0]]
    c.value, c.font, c.fill, c.alignment = text, F_BODY, FILL_NOTE, WRAP

def color_bar(chart, color):
    for s in chart.series:
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color

def labels(chart, pct=True):
    for s in chart.series:
        s.dLbls = DataLabelList(); s.dLbls.showVal = True
        s.dLbls.showSerName = s.dLbls.showCatName = s.dLbls.showLegendKey = False
        if pct: s.dLbls.numFmt = "0%"

# ---------------------------------------------------------------- Clean Data tab
COLS = ["respondent_id", "collection_date", "level", "faculty", "department", "dept_group", "primary_provider",
        "connection_type", "freq_use", "freq_use_code", "access_location", "rating", "rating_code", "poor_freq",
        "poor_freq_code", "disconnect", "disconnect_code", "satisfaction", "satisfaction_code", "impact", "impact_code",
        "high_impact", "missed_deadline", "weekly_spend", "weekly_spend_code", "cei", "conn_group",
        "prov_mtn", "prov_glo", "prov_airtel", "prov_9mobile",
        "time_morning_6_am_11_am", "time_afternoon_12_pm_3_pm", "time_evening_4_pm_7_pm", "time_night_8_pm_onwards",
        "time_no_noticeable_difference",
        "act_attending_online_classes", "act_online_research", "act_downloading_lecture_materials",
        "act_watching_educational_videos", "act_submitting_assignments", "act_online_examinations",
        "uni_install_more_wi_fi_hotspots", "uni_improve_existing_network_infrastructure",
        "uni_partner_with_telecommunications_providers", "uni_provide_affordable_internet_access", "uni_other",
        "effect_type", "challenge_text", "suggestion_text"]
data = df[COLS].copy()
N = len(data); R0, R1 = 2, N + 1
DS = "'Clean Data'"
L = {c: get_column_letter(i + 1) for i, c in enumerate(COLS)}
def rng(col): return f"{DS}!${L[col]}${R0}:${L[col]}${R1}"

wd = wb.create_sheet("Clean Data")
wd.sheet_view.showGridLines = False
for j, c in enumerate(COLS, start=1):
    cell = wd.cell(row=1, column=j, value=c)
    cell.font, cell.fill = F_HEAD, FILL_HEAD
    wd.column_dimensions[get_column_letter(j)].width = 14 if not c.endswith("_text") else 50
for i, row in enumerate(data.itertuples(index=False), start=2):
    for j, v in enumerate(row, start=1):
        if pd.notna(v):
            wd.cell(row=i, column=j, value=v.item() if hasattr(v, "item") else v).font = F_BODY
wd.freeze_panes = "B2"
wd.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{R1}"
wd["A1"].comment = Comment(
    "Source: NSUK internet connectivity survey (Google Form 'Untitled form (Responses)'), 84 rows, "
    "collected 18-22 Nov 2026. Cleaned in notebooks/NSUK_Internet_Connectivity_Analysis.ipynb; "
    "file data/processed/survey_clean.csv. Derived columns (_code, cei, conn_group, high_impact, 0/1 "
    "indicators) are defined on the Data Dictionary tab.", "Analyst")

# ---------------------------------------------------------------- Overview
ws = sheet("Overview", "Internet Connectivity and Students' Learning Experience at NSUK",
           "NS-DTS 223 Group Practical Assignment, Topic 2.  Online survey of NSUK students, 18-22 November 2026.",
           [52, 16, 3, 22, 40, 70])
h2(ws, "A4", "Central question")
note_box(ws, "A5:F5", "How do students experience internet connectivity across the NSUK campus, and how is a poorer "
         "connectivity experience associated with disruption to their learning?")
ws.row_dimensions[5].height = 32
h2(ws, "A7", "Headline figures")
header(ws, 8, ["Measure", "Value"])
kpis = [
    ("Valid responses", f"=COUNTA({rng('respondent_id')})", "0"),
    ("Use the internet very often or always for academic work", f"=COUNTIF({rng('freq_use_code')},\">=4\")/COUNT({rng('freq_use_code')})", PCT),
    ("Rely mainly on their own mobile data", f"=COUNTIF({rng('connection_type')},\"Mobile data\")/B9", PCT),
    ("Mainly use campus Wi-Fi", f"=COUNTIF({rng('connection_type')},\"Campus Wi-Fi\")/B9", PCT),
    ("Rate campus connection Poor or Very poor", f"=COUNTIF({rng('rating_code')},\"<=2\")/COUNT({rng('rating_code')})", PCT),
    ("Experience poor connectivity often or always", f"=COUNTIF({rng('poor_freq_code')},\">=4\")/COUNT({rng('poor_freq_code')})", PCT),
    ("Disconnected unexpectedly often or always", f"=COUNTIF({rng('disconnect_code')},\">=4\")/COUNT({rng('disconnect_code')})", PCT),
    ("Satisfied or very satisfied with provider speed", f"=COUNTIF({rng('satisfaction_code')},\">=4\")/COUNT({rng('satisfaction_code')})", PCT),
    ("Academic performance highly or very highly affected", f"=AVERAGE({rng('high_impact')})", PCT),
    ("Have missed a deadline because of poor connectivity", f"=COUNTIF({rng('missed_deadline')},\"Yes\")/COUNTA({rng('missed_deadline')})", PCT),
    ("Spend more than N3,000 a week on data", f"=COUNTIF({rng('weekly_spend_code')},\">=4\")/COUNT({rng('weekly_spend_code')})", PCT),
    ("Median Connectivity Experience Index (0 worst, 100 best)", f"=MEDIAN({rng('cei')})", "0.0"),
]
for i, (lab, f, fmt) in enumerate(kpis, start=9):
    ws.cell(row=i, column=1, value=lab); ws.cell(row=i, column=2, value=f)
body(ws, 9, 8 + len(kpis), 1, 2)
for i, (_, _, fmt) in enumerate(kpis, start=9):
    ws.cell(row=i, column=2).number_format = fmt; ws.cell(row=i, column=2).font = F_KPI
h2(ws, "D7", "The data story")
header(ws, 8, ["Stage", "Question", "Finding"], col=4)
story = [
    ("1. Overview", "How do students get online?", "Most students rely on their own mobile data; very few use campus Wi-Fi. MTN is the most used network. (tab 1 Access)"),
    ("2. Distribution", "Where is the experience worse?", "The experience is middling in every location. Location differences are small and not statistically significant, so the problem is campus-wide. (tab 2 Location)"),
    ("3. Pattern", "When is connectivity best?", "Night is reported best and the afternoon worst, consistent with congestion during teaching hours. (tab 3 Time of Day)"),
    ("4. Relationship", "Is poorer connectivity linked to academic impact?", "Students who rate their connection poorly report greater academic impact (Spearman rho = -0.28, p = 0.009). (tab 4 Impact)"),
    ("5. Key insight", "Who misses deadlines?", "Poor-connectivity students are 2.6 times as likely to report high impact. Most students in every group have missed a deadline. (tab 5 Deadlines)"),
    ("6. Action", "What should the university do?", "Campus-wide Wi-Fi in academic spaces, operator capacity agreements, outage-tolerant teaching, cheaper access, regular speed testing. (tab 6 Recommendations)"),
]
for i, row in enumerate(story, start=9):
    for j, v in enumerate(row):
        ws.cell(row=i, column=4 + j, value=v)
body(ws, 9, 8 + len(story), 4, 6)
for r in range(9, 9 + len(story)):
    ws.row_dimensions[r].height = 42
    for c in (4, 5, 6):
        ws.cell(row=r, column=c).alignment = WRAP
    ws.cell(row=r, column=4).font = F_BOLD
h2(ws, "A23", "About the data")
note_box(ws, "A24:F24", "84 responses from a convenience sample (79% Natural and Applied Sciences, 87% 200-300 Level), "
         "so results are indicative rather than campus-wide estimates. All measures are self-reported; no speed tests "
         "were run. Every figure on the summary tabs is a formula over the Clean Data tab. Statistical tests come from "
         "the Python notebook (Statistical Tests tab).")
ws.row_dimensions[24].height = 58
ws.freeze_panes = "A4"

# ---------------------------------------------------------------- 1 Access
ws = sheet("1 Access", "Most students rely on their own mobile data to get online",
           "Visualization 1, Overview.  Providers is a multi-select question, so its percentages can add up to more than 100%.",
           [34, 14, 14, 4, 30, 14, 14])
h2(ws, "A4", "Network providers used")
header(ws, 5, ["Provider", "Students", "% of respondents"])
for i, (lab, col) in enumerate([("MTN", "prov_mtn"), ("Glo", "prov_glo"), ("Airtel", "prov_airtel"), ("9mobile", "prov_9mobile")], start=6):
    ws.cell(row=i, column=1, value=lab)
    ws.cell(row=i, column=2, value=f"=SUM({rng(col)})")
    ws.cell(row=i, column=3, value=f"=B{i}/Overview!$B$9")
body(ws, 6, 9, 1, 3, {3: PCT})
ws["A10"] = "Students using more than one network"; ws["B10"] = f"=COUNTIF({rng('primary_provider')},\"Multiple SIMs\")"
body(ws, 10, 10, 1, 2, band=False)
h2(ws, "E4", "Main type of connection")
header(ws, 5, ["Connection type", "Students", "% of respondents"], col=5)
for i, lab in enumerate(["Mobile data", "Campus Wi-Fi", "Personal Wi-Fi"], start=6):
    ws.cell(row=i, column=5, value=lab)
    ws.cell(row=i, column=6, value=f"=COUNTIF({rng('connection_type')},E{i})")
    ws.cell(row=i, column=7, value=f"=F{i}/SUM($F$6:$F$8)")
body(ws, 6, 8, 5, 7, {7: PCT})
h2(ws, "E10", "Frequency of academic internet use")
header(ws, 11, ["Frequency", "Students", "% of valid answers"], col=5)
for i, lab in enumerate(["Always", "Very often", "Sometimes", "Rarely", "Never"], start=12):
    ws.cell(row=i, column=5, value=lab)
    ws.cell(row=i, column=6, value=f"=COUNTIF({rng('freq_use')},E{i})")
    ws.cell(row=i, column=7, value=f"=F{i}/SUM($F$12:$F$16)")
body(ws, 12, 16, 5, 7, {7: PCT})
ch = BarChart(); ch.type = "bar"; ch.title = "Network providers used (% of respondents)"
ch.add_data(Reference(ws, min_col=3, min_row=5, max_row=9), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
ch.y_axis.numFmt = "0%"; ch.y_axis.majorGridlines = None; ch.legend = None; ch.x_axis.scaling.orientation = "maxMin"
color_bar(ch, BLUE); labels(ch); ch.height, ch.width = 7.5, 14
ws.add_chart(ch, "A19")
ch2 = BarChart(); ch2.type = "col"; ch2.title = "Main type of connection (% of respondents)"
ch2.add_data(Reference(ws, min_col=7, min_row=5, max_row=8), titles_from_data=True)
ch2.set_categories(Reference(ws, min_col=5, min_row=6, max_row=8))
ch2.y_axis.numFmt = "0%"; ch2.legend = None; color_bar(ch2, ORANGE); labels(ch2); ch2.height, ch2.width = 7.5, 13
ws.add_chart(ch2, "E19")
note_box(ws, "A35:G35", "Reading: 82% of students rely mainly on their own mobile data and only 10% mainly use campus Wi-Fi, "
         "while 89% use the internet very often or always for academic work. Students' learning depends on commercial "
         "networks they pay for themselves.")
ws.row_dimensions[35].height = 45

# ---------------------------------------------------------------- 2 Location
ws = sheet("2 Location", "No campus location stands out: the experience is middling everywhere",
           "Visualization 2, Distribution.  Connectivity Experience Index (CEI): 0 = worst, 100 = best; mean of the rating, "
           "reversed poor-connectivity and disconnection frequency, and satisfaction.",
           [26, 11, 13, 13, 13, 13, 4, 22, 11, 13, 13])
h2(ws, "A4", "CEI by main access location")
header(ws, 5, ["Location", "Students", "Median CEI", "Mean CEI", "% rating Poor", "% rating Good"])
locs = ["Faculty buildings", "Library", "Lecture halls", "Hostels", "Open spaces", "Other"]
for i, lab in enumerate(locs, start=6):
    ws.cell(row=i, column=1, value=lab)
    ws.cell(row=i, column=2, value=f"=COUNTIF({rng('access_location')},A{i})")
    ws.cell(row=i, column=3, value=f"=MEDIAN(FILTER({rng('cei')},{rng('access_location')}=A{i}))")
    ws.cell(row=i, column=4, value=f"=AVERAGEIF({rng('access_location')},A{i},{rng('cei')})")
    ws.cell(row=i, column=5, value=f"=COUNTIFS({rng('access_location')},A{i},{rng('conn_group')},\"Poor\")/B{i}")
    ws.cell(row=i, column=6, value=f"=COUNTIFS({rng('access_location')},A{i},{rng('conn_group')},\"Good\")/B{i}")
r = 6 + len(locs)
ws.cell(row=r, column=1, value="All students").font = F_BOLD
ws.cell(row=r, column=2, value=f"=SUM(B6:B{r-1})")
ws.cell(row=r, column=3, value=f"=MEDIAN({rng('cei')})"); ws.cell(row=r, column=4, value=f"=AVERAGE({rng('cei')})")
ws.cell(row=r, column=5, value=f"=COUNTIF({rng('conn_group')},\"Poor\")/COUNTA({rng('conn_group')})")
ws.cell(row=r, column=6, value=f"=COUNTIF({rng('conn_group')},\"Good\")/COUNTA({rng('conn_group')})")
body(ws, 6, r, 1, 6, {3: "0.0", 4: "0.0", 5: PCT, 6: PCT})
for c in range(1, 7): ws.cell(row=r, column=c).font = F_BOLD
ws.cell(row=r + 1, column=1, value="One respondent did not state a location.").font = F_SUB
h2(ws, "H4", "CEI by network provider")
header(ws, 5, ["Provider", "Students", "Median CEI", "Mean CEI"], col=8)
for i, lab in enumerate(["MTN", "Glo", "Airtel", "Multiple SIMs"], start=6):
    ws.cell(row=i, column=8, value=lab)
    ws.cell(row=i, column=9, value=f"=COUNTIF({rng('primary_provider')},H{i})")
    ws.cell(row=i, column=10, value=f"=MEDIAN(FILTER({rng('cei')},{rng('primary_provider')}=H{i}))")
    ws.cell(row=i, column=11, value=f"=AVERAGEIF({rng('primary_provider')},H{i},{rng('cei')})")
body(ws, 6, 9, 8, 11, {10: "0.0", 11: "0.0"})
ch = BarChart(); ch.type = "bar"; ch.title = "Median CEI by access location (0-100)"
ch.add_data(Reference(ws, min_col=3, min_row=5, max_row=5 + len(locs)), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=5 + len(locs)))
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 100; ch.legend = None; ch.x_axis.scaling.orientation = "maxMin"
color_bar(ch, BLUE); labels(ch, pct=False); ch.height, ch.width = 8, 16
ws.add_chart(ch, "A16")
note_box(ws, "A33:K33", "Reading: median CEI ranges from 44 to 56 across locations, around or below the midpoint of 50. "
         "Differences between locations (Kruskal-Wallis p = 0.60) and providers (p = 0.24) are not statistically "
         "significant, and the spread within each location is much wider than the gap between them. Students' comments "
         "single out faculty buildings and the FNAS lecture theatres (LT1, LT2). The full box plot is "
         "reports/figures/story_02_cei_by_location.png.")
ws.row_dimensions[33].height = 60

# ---------------------------------------------------------------- 3 Time of Day
ws = sheet("3 Time of Day", "Connectivity is reported best at night, when campus demand is lowest",
           "Visualization 3, Pattern.  Multi-select question; % of students who named at least one time slot.",
           [30, 12, 22, 4, 60])
h2(ws, "A4", "Best time for internet connection")
header(ws, 5, ["Time slot", "Students", "% of students naming a slot"])
slots = [("Morning (6 AM-11 AM)", "time_morning_6_am_11_am"), ("Afternoon (12 PM-3 PM)", "time_afternoon_12_pm_3_pm"),
         ("Evening (4 PM-7 PM)", "time_evening_4_pm_7_pm"), ("Night (8 PM onwards)", "time_night_8_pm_onwards")]
for i, (lab, col) in enumerate(slots, start=6):
    ws.cell(row=i, column=1, value=lab); ws.cell(row=i, column=2, value=f"=SUM({rng(col)})")
    ws.cell(row=i, column=3, value=f"=B{i}/$B$12")
body(ws, 6, 9, 1, 3, {3: PCT})
ws["A11"] = "Reported no noticeable difference"; ws["B11"] = f"=SUM({rng('time_no_noticeable_difference')})"
ws["A12"] = "Students naming at least one slot"; ws["B12"] = "=Overview!B9-B11"
body(ws, 11, 12, 1, 2, band=False)
lc = LineChart(); lc.title = "Students naming each slot as best (%)"
lc.add_data(Reference(ws, min_col=3, min_row=5, max_row=9), titles_from_data=True)
lc.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
lc.y_axis.numFmt = "0%"; lc.y_axis.scaling.min = 0; lc.legend = None
s = lc.series[0]; s.graphicalProperties.line.solidFill = BLUE; s.graphicalProperties.line.width = 28575
s.marker.symbol = "circle"; s.marker.size = 8; s.marker.graphicalProperties.solidFill = BLUE
labels(lc); lc.height, lc.width = 8, 16
ws.add_chart(lc, "A15")
note_box(ws, "E5:E9", "Reading: 45% of students who named a time say the connection is best after 8 PM, compared with 19% "
         "for the afternoon. This fits congestion during teaching hours, when most students and staff are online at once. "
         "The best connectivity comes when students are least able to attend classes or use campus facilities.")

# ---------------------------------------------------------------- 4 Impact
ws = sheet("4 Impact", "Students who rate their connection poorly report a greater academic impact",
           "Visualization 4, Relationship.  Rows: rating of campus connection (Q8).  Columns: effect of poor connectivity "
           "on academic performance (Q15).",
           [22, 12, 12, 13, 12, 13, 10])
RAT = ["Very poor", "Poor", "Average", "Good", "Excellent"]
IMP = ["Not at all", "Slightly", "Moderately", "Highly", "Very highly"]
h2(ws, "A4", "Number of students")
header(ws, 5, ["Rating \\ Impact"] + IMP + ["Total"])
for i, rt in enumerate(RAT, start=6):
    ws.cell(row=i, column=1, value=rt)
    for j in range(5):
        col = get_column_letter(2 + j)
        ws.cell(row=i, column=2 + j, value=f"=COUNTIFS({rng('rating')},$A{i},{rng('impact')},{col}$5)")
    ws.cell(row=i, column=7, value=f"=SUM(B{i}:F{i})")
body(ws, 6, 10, 1, 7)
h2(ws, "A13", "Row percentages (heatmap)")
header(ws, 14, ["Rating \\ Impact"] + IMP + ["Total"])
for i, rt in enumerate(RAT, start=15):
    ws.cell(row=i, column=1, value=rt)
    for j in range(5):
        col = get_column_letter(2 + j)
        ws.cell(row=i, column=2 + j, value=f"=IF($G{i-9}=0,0,{col}{i-9}/$G{i-9})")
    ws.cell(row=i, column=7, value=f"=SUM(B{i}:F{i})")
body(ws, 15, 19, 1, 7, {c: PCT for c in range(2, 8)}, band=False)
ws.conditional_formatting.add("B15:F19", ColorScaleRule(start_type="num", start_value=0, start_color="F7F7F5",
                                                        end_type="num", end_value=0.6, end_color="256ABF"))
h2(ws, "A22", "Association")
header(ws, 23, ["Measure", "Value"])
ws["A24"] = "Pearson r, rating code vs impact code"; ws["B24"] = f"=CORREL({rng('rating_code')},{rng('impact_code')})"
ws["A25"] = "Pearson r, CEI vs impact code"; ws["B25"] = f"=CORREL({rng('cei')},{rng('impact_code')})"
ws["A26"] = "Spearman rho, rating vs impact (notebook)"; ws["B26"] = "='Statistical Tests'!C5"
ws["A27"] = "Spearman rho, CEI vs impact (notebook)"; ws["B27"] = "='Statistical Tests'!C6"
body(ws, 24, 27, 1, 2, {2: "0.00"})
ws.column_dimensions["A"].width = 40
note_box(ws, "A30:G30", "Reading: the darker cells run from top right to bottom left. Half of students rating their connection "
         "Very poor say their performance is very highly affected; none of those rating it Good or Excellent do. The "
         "association is negative and statistically significant, and it holds after controlling for data spend, level "
         "and frequency of use (partial rho = -0.28, p = 0.011). It is an association, not proof of cause.")
ws.row_dimensions[30].height = 60

# ---------------------------------------------------------------- 5 Deadlines
ws = sheet("5 Deadlines", "Poor-connectivity students are 2.6 times as likely to report high academic impact",
           "Visualization 5, Key insight.  Connectivity group from Q8: Poor = Very poor or Poor; Good = Good or Excellent.",
           [26, 12, 22, 22, 4, 50])
header(ws, 5, ["Connectivity group", "Students", "Missed a deadline", "High academic impact"])
for i, g in enumerate(["Poor", "Average", "Good"], start=6):
    ws.cell(row=i, column=1, value=g)
    ws.cell(row=i, column=2, value=f"=COUNTIF({rng('conn_group')},A{i})")
    ws.cell(row=i, column=3, value=f"=COUNTIFS({rng('conn_group')},A{i},{rng('missed_deadline')},\"Yes\")/COUNTIFS({rng('conn_group')},A{i},{rng('missed_deadline')},\"<>\")")
    ws.cell(row=i, column=4, value=f"=AVERAGEIF({rng('conn_group')},A{i},{rng('high_impact')})")
ws["A9"] = "All students"; ws["B9"] = "=SUM(B6:B8)"; ws["C9"] = "=Overview!B18"; ws["D9"] = "=Overview!B17"
body(ws, 6, 9, 1, 4, {3: PCT, 4: PCT})
for c in range(1, 5): ws.cell(row=9, column=c).font = F_BOLD
ws["A11"] = "Ratio of high-impact rates, Poor vs Average and Good"
ws["C11"] = f"=D6/(COUNTIFS({rng('conn_group')},\"<>Poor\",{rng('high_impact')},1)/COUNTIFS({rng('conn_group')},\"<>Poor\",{rng('high_impact')},\"<>\"))"
ws["C11"].number_format = '0.0"x"'; ws["A11"].font = F_BOLD; ws["C11"].font = F_KPI
cc = BarChart(); cc.type = "col"; cc.grouping = "clustered"; cc.title = "Missed deadline and high academic impact by connectivity group"
cc.add_data(Reference(ws, min_col=3, max_col=4, min_row=5, max_row=8), titles_from_data=True)
cc.set_categories(Reference(ws, min_col=1, min_row=6, max_row=8))
cc.y_axis.numFmt = "0%"; cc.y_axis.scaling.min, cc.y_axis.scaling.max = 0, 1
cc.series[0].graphicalProperties.solidFill = ORANGE; cc.series[0].graphicalProperties.line.solidFill = ORANGE
cc.series[1].graphicalProperties.solidFill = BLUE; cc.series[1].graphicalProperties.line.solidFill = BLUE
labels(cc); cc.legend.position = "b"; cc.height, cc.width = 8.5, 16
ws.add_chart(cc, "A14")
note_box(ws, "F5:F11", "Reading: 65% of students with a Poor connection say their academic performance is highly affected, "
         "against about 25% in the other groups (chi-square p = 0.005). Missed deadlines are common in every group: "
         "85% Poor, 72% Average, 59% Good. That trend is in the expected direction but is not statistically "
         "significant with 84 responses (p = 0.15).")

# ---------------------------------------------------------------- 6 Recommendations
ws = sheet("6 Recommendations", "Students want more campus Wi-Fi and better infrastructure",
           "Supporting evidence for the recommendations.  Multi-select questions; % of respondents.",
           [44, 12, 16, 4, 40, 12])
h2(ws, "A4", "What the university should do (Q18)")
header(ws, 5, ["Action", "Students", "% of respondents"])
uni = [("Install more Wi-Fi hotspots", "uni_install_more_wi_fi_hotspots"),
       ("Improve existing network infrastructure", "uni_improve_existing_network_infrastructure"),
       ("Partner with telecommunications providers", "uni_partner_with_telecommunications_providers"),
       ("Provide affordable internet access", "uni_provide_affordable_internet_access"), ("Other", "uni_other")]
for i, (lab, col) in enumerate(uni, start=6):
    ws.cell(row=i, column=1, value=lab); ws.cell(row=i, column=2, value=f"=SUM({rng(col)})")
    ws.cell(row=i, column=3, value=f"=B{i}/Overview!$B$9")
body(ws, 6, 10, 1, 3, {3: PCT})
h2(ws, "E4", "Activities most affected (Q13)")
header(ws, 5, ["Activity", "Students", "% of respondents"], col=5)
acts = [("Attending online classes", "act_attending_online_classes"), ("Online research", "act_online_research"),
        ("Downloading lecture materials", "act_downloading_lecture_materials"),
        ("Watching educational videos", "act_watching_educational_videos"),
        ("Submitting assignments", "act_submitting_assignments"), ("Online examinations", "act_online_examinations")]
for i, (lab, col) in enumerate(acts, start=6):
    ws.cell(row=i, column=5, value=lab); ws.cell(row=i, column=6, value=f"=SUM({rng(col)})")
    ws.cell(row=i, column=7, value=f"=F{i}/COUNT({rng(col)})")
body(ws, 6, 11, 5, 7, {7: PCT})
ws.column_dimensions["G"].width = 16
ch = BarChart(); ch.type = "bar"; ch.title = "What the university should do (% of respondents)"
ch.add_data(Reference(ws, min_col=3, min_row=5, max_row=9), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
ch.y_axis.numFmt = "0%"; ch.legend = None; ch.x_axis.scaling.orientation = "maxMin"
color_bar(ch, BLUE); labels(ch); ch.height, ch.width = 7.5, 16
ws.add_chart(ch, "A13")
h2(ws, "A29", "Recommendations")
header(ws, 30, ["Recommendation", "For", "Evidence"])
ws.merge_cells("C30:G30")
recs = [
    ("1. Expand campus Wi-Fi as a campus-wide service, starting with faculty buildings, FNAS LT1 and LT2, the library and lecture halls, sized for daytime peaks.",
     "Management, ICT", "Only 10% use campus Wi-Fi; 50% ask for more hotspots; the experience is middling in every location."),
    ("2. Agree capacity upgrades with MTN, Glo and Airtel for the campus during teaching hours.",
     "Management, operators", "52% face poor connectivity often or always; night is reported best; disconnections are the strongest correlate of academic impact."),
    ("3. Make teaching tolerant of outages: record online classes, provide downloadable materials, allow grace periods for outage-related late submissions.",
     "Lecturers, Senate", "71% have missed a deadline; online classes (49%) and research (46%) are the most affected activities."),
    ("4. Cut students' data costs through zero-rated portal, LMS and e-library access and student bundles.",
     "Management, SU, operators", "65% spend more than N3,000 a week; 82% depend on their own data."),
    ("5. Run scheduled speed tests at fixed campus points each semester and publish the results.",
     "ICT directorate", "This study relies on perceptions; measured data would show where and when to invest."),
]
for i, (a, b, c) in enumerate(recs, start=31):
    ws.cell(row=i, column=1, value=a); ws.cell(row=i, column=2, value=b); ws.cell(row=i, column=3, value=c)
    ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=7)
body(ws, 31, 35, 1, 7)
for r in range(31, 36):
    ws.row_dimensions[r].height = 48
    for c in (1, 2, 3): ws.cell(row=r, column=c).alignment = WRAP

# ---------------------------------------------------------------- Statistical Tests
ws = sheet("Statistical Tests", "Statistical tests behind the findings",
           "Spreadsheets cannot reproduce these tests, so the results are imported from the Python notebook.",
           [46, 34, 12, 12, 40])
header(ws, 4, ["Test", "Variables", "Statistic", "p-value", "Interpretation"])
tests = [
    ("Spearman rank correlation", "Rating (Q8) vs academic impact (Q15)", -0.284, 0.009, "Significant, weak-to-moderate negative"),
    ("Spearman rank correlation", "CEI vs academic impact (Q15)", -0.307, 0.005, "Significant, moderate negative"),
    ("Spearman rank correlation", "Disconnection frequency vs academic impact", 0.364, 0.001, "Strongest single correlate of impact"),
    ("Partial Spearman (controls: spend, level, use)", "CEI vs academic impact", -0.28, 0.011, "Association holds after adjustment"),
    ("Chi-square, Cramer's V = 0.36", "Connectivity group vs high academic impact", 10.48, 0.005, "Significant"),
    ("Chi-square, Cramer's V = 0.21", "Connectivity group vs missed deadline", 3.74, 0.154, "Not significant"),
    ("Fisher's exact, odds ratio", "Poor vs other groups, missed deadline", 2.83, 0.159, "Not significant"),
    ("Kruskal-Wallis H", "CEI by access location", 3.65, 0.600, "Not significant"),
    ("Kruskal-Wallis H", "CEI by primary provider", 4.24, 0.236, "Not significant"),
    ("Kruskal-Wallis H", "CEI by connection type", 1.06, 0.590, "Not significant"),
    ("Kruskal-Wallis H", "CEI by level (100 and 500 Level excluded, n < 5)", 4.10, 0.128, "Not significant"),
    ("Mann-Whitney U, rank-biserial r = -0.27", "CEI, NAS vs other faculties", 756, 0.076, "Borderline; treat as a hypothesis"),
    ("Kruskal-Wallis H", "CEI by weekly data spend", 9.17, 0.057, "Borderline; treat as a hypothesis"),
    ("Cronbach's alpha", "Four CEI items", 0.78, None, "Acceptable internal consistency"),
]
for i, row in enumerate(tests, start=5):
    for j, v in enumerate(row, start=1):
        if v is not None: ws.cell(row=i, column=j, value=v)
body(ws, 5, 4 + len(tests), 1, 5, {3: "0.00", 4: "0.000"})
for r in range(5, 5 + len(tests)):
    for c in (1, 2, 5): ws.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws["A4"].comment = Comment("Source: notebooks/NSUK_Internet_Connectivity_Analysis.ipynb, sections 6.2 and 9.1-9.4; "
                           "also reports/tables/subgroup_tests_cei.csv and spearman_matrix.csv. Run 22 Nov 2026.", "Analyst")

# ---------------------------------------------------------------- Data Quality
ws = sheet("Data Quality", "Data quality audit and cleaning log",
           "84 responses kept; no duplicates; every change is listed below.", [30, 16, 10, 70, 60])
audit = pd.read_csv(TAB / "data_quality_audit.csv")
h2(ws, "A4", "Issues found before cleaning")
header(ws, 5, ["Check", "Column", "Rows", "Finding", "Action"])
for i, row in enumerate(audit.itertuples(index=False), start=6):
    for j, v in enumerate(row, start=1): ws.cell(row=i, column=j, value=v)
end = 5 + len(audit)
body(ws, 6, end, 1, 5)
for r in range(6, end + 1):
    for c in (1, 4, 5): ws.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws["A5"].comment = Comment("Source: reports/tables/data_quality_audit.csv, generated by the notebook (section 4).", "Analyst")
log = pd.read_csv(TAB / "cleaning_log.csv")[["stage", "column", "rows_affected", "action", "rationale"]]
s = end + 3
h2(ws, f"A{s}", "Cleaning log")
header(ws, s + 1, ["Stage", "Column", "Rows", "Action", "Rationale"])
for i, row in enumerate(log.itertuples(index=False), start=s + 2):
    for j, v in enumerate(row, start=1): ws.cell(row=i, column=j, value=v)
body(ws, s + 2, s + 1 + len(log), 1, 5)
for r in range(s + 2, s + 2 + len(log)):
    for c in (1, 2, 4, 5): ws.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws[f"A{s+1}"].comment = Comment("Source: reports/tables/cleaning_log.csv, generated by the notebook (section 5).", "Analyst")

# ---------------------------------------------------------------- Data Dictionary
ws = sheet("Data Dictionary", "Data dictionary for the Clean Data tab", "Columns shown in the Clean Data tab.",
           [40, 55, 18, 60])
dd = pd.read_csv(TAB / "data_dictionary.csv").set_index("variable").reindex(COLS).reset_index()
header(ws, 4, ["Variable", "Description", "Type", "Values / coding"])
for i, row in enumerate(dd[["variable", "description", "type", "values / coding"]].itertuples(index=False), start=5):
    for j, v in enumerate(row, start=1):
        if pd.notna(v): ws.cell(row=i, column=j, value=v)
body(ws, 5, 4 + len(dd), 1, 4)
for r in range(5, 5 + len(dd)):
    for c in (2, 3, 4): ws.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

# move Clean Data to the end
wb.move_sheet("Clean Data", offset=len(wb.sheetnames))
for w in wb.worksheets:
    w.sheet_properties.tabColor = {"Overview": NAVY, "Clean Data": "8A8984", "Statistical Tests": AQUA,
                                   "Data Quality": AQUA, "Data Dictionary": AQUA}.get(w.title, BLUE)
wb.save(OUT)
print("saved", OUT, OUT.stat().st_size, "bytes;", wb.sheetnames)
