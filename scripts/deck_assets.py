"""Compute the numbers and the one image the slide deck needs.

Writes scripts/deck/deck_data.json and scripts/deck/boxplot_location.png from
data/processed/survey_clean.csv. Run after the notebook, before build_deck.js.
"""
import json
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "scripts/deck"; OUT.mkdir(exist_ok=True)
df = pd.read_csv(ROOT / "data/processed/survey_clean.csv")
N = len(df)
pct = lambda s: round(float(s) * 100, 1)

RAT = ["Very poor", "Poor", "Average", "Good", "Excellent"]
IMP = ["Not at all", "Slightly", "Moderately", "Highly", "Very highly"]
ct = pd.crosstab(df.rating, df.impact).reindex(index=RAT, columns=IMP, fill_value=0)
rowpct = ct.div(ct.sum(axis=1), axis=0)
grp = (df.groupby("conn_group").agg(n=("respondent_id", "size"), missed=("missed_deadline_bin", "mean"),
                                    high=("high_impact", "mean")).reindex(["Poor", "Average", "Good"]))
slots = {"Morning": "time_morning_6_am_11_am", "Afternoon": "time_afternoon_12_pm_3_pm",
         "Evening": "time_evening_4_pm_7_pm", "Night": "time_night_8_pm_onwards"}
named = df[list(slots.values())].sum(axis=1) > 0
uni = {"Install more Wi-Fi hotspots": "uni_install_more_wi_fi_hotspots",
       "Improve existing infrastructure": "uni_improve_existing_network_infrastructure",
       "Partner with telecom providers": "uni_partner_with_telecommunications_providers",
       "Provide affordable access": "uni_provide_affordable_internet_access"}

data = {
    "n": N,
    "level": df.level.value_counts().reindex(["100 Level", "200 Level", "300 Level", "400 Level", "500 Level"]).fillna(0).astype(int).to_dict(),
    "faculty": {"Natural & Applied Sciences": int((df.faculty == "Natural and Applied Sciences").sum()),
                "Other faculties": int((df.faculty != "Natural and Applied Sciences").sum())},
    "providers": {k: pct(df[c].mean()) for k, c in [("MTN", "prov_mtn"), ("Glo", "prov_glo"), ("Airtel", "prov_airtel"), ("9mobile", "prov_9mobile")]},
    "conn_type": {k: pct((df.connection_type == k).mean()) for k in ["Mobile data", "Campus Wi-Fi", "Personal Wi-Fi"]},
    "kpi": {
        "academic_use": pct((df.freq_use_code >= 4).sum() / df.freq_use_code.notna().sum()),
        "poor_often": pct((df.poor_freq_code >= 4).mean()),
        "disconnect_often": pct((df.disconnect_code >= 4).sum() / df.disconnect_code.notna().sum()),
        "satisfied": pct((df.satisfaction_code >= 4).sum() / df.satisfaction_code.notna().sum()),
        "rate_poor": pct((df.rating_code <= 2).mean()),
        "high_impact": pct(df.high_impact.mean()),
        "missed": pct(df.missed_deadline_bin.mean()),
        "spend_3k": pct((df.weekly_spend_code >= 4).mean()),
        "median_cei": float(df.cei.median()),
    },
    "time": {k: pct(df.loc[named, c].mean()) for k, c in slots.items()},
    "time_named": int(named.sum()), "time_nodiff": int(df.time_no_noticeable_difference.sum()),
    "heat_counts": ct.values.tolist(), "heat_pct": rowpct.round(3).values.tolist(),
    "heat_rows": RAT, "heat_cols": IMP, "heat_n": ct.sum(axis=1).tolist(),
    "groups": {g: {"n": int(r.n), "missed": pct(r.missed), "high": pct(r.high)} for g, r in grp.iterrows()},
    "uni": {k: pct(df[c].mean()) for k, c in uni.items()},
}
(OUT / "deck_data.json").write_text(json.dumps(data, indent=1), encoding="utf-8")

# ---- box plot (no native slide chart type), single-hue teal, sized for the slide
SEQ = ["#E6F4F6", "#B9E2E8", "#6CC1CF", "#2A9DB0", "#137C8F", "#0E5A6B", "#0B3C49"]
INK, INK2, GRID = "#1B2B30", "#5B6B70", "#E3E9EA"
plt.rcParams.update({"font.family": "Calibri" if "Calibri" in {f.name for f in matplotlib.font_manager.fontManager.ttflist} else "DejaVu Sans",
                     "font.size": 15, "axes.edgecolor": GRID, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.spines.left": False, "xtick.color": INK2, "ytick.color": INK, "axes.grid": True,
                     "grid.color": GRID, "axes.axisbelow": True})
order = df.groupby("access_location")["cei"].median().sort_values().index.tolist()
vals = [df.loc[df.access_location == o, "cei"].dropna() for o in order]
fig, ax = plt.subplots(figsize=(8.2, 4.9), dpi=220)
ax.boxplot(vals, orientation="horizontal", widths=0.56, patch_artist=True, showfliers=False,
           medianprops=dict(color=SEQ[6], lw=3), boxprops=dict(facecolor=SEQ[1], edgecolor=SEQ[4], lw=1.4),
           whiskerprops=dict(color=SEQ[4], lw=1.4), capprops=dict(color=SEQ[4], lw=1.4))
rng = np.random.default_rng(7)
for i, v in enumerate(vals, start=1):
    ax.scatter(v, i + rng.uniform(-.15, .15, len(v)), s=30, color=SEQ[5], alpha=.8, edgecolor="white", linewidth=.8, zorder=3)
ax.set_yticks(range(1, len(order) + 1), [f"{o}  ({len(v)})" for o, v in zip(order, vals)])
med = df.cei.median()
ax.axvline(med, color=SEQ[6], lw=1.2, ls=(0, (1, 0)), alpha=.5)
ax.text(med + 1, len(order) + .55, f"Overall median {med:.0f}", color=INK2, fontsize=13)
ax.set_xlim(0, 100); ax.set_ylim(.4, len(order) + .8)
ax.set_xlabel("Connectivity Experience Index  (0 = worst, 100 = best)", color=INK2, fontsize=14)
ax.grid(axis="y", visible=False); ax.tick_params(axis="y", length=0)
fig.tight_layout(); fig.savefig(OUT / "boxplot_location.png", facecolor="white"); plt.close(fig)
print("wrote", OUT / "deck_data.json", "and boxplot_location.png")


# ---- slide charts as images (Google Slides drops category labels from imported native charts)
plt.rcParams.update({"axes.spines.left": False, "axes.spines.bottom": True, "font.size": 14})
def _fig(w, h):
    fig, ax = plt.subplots(figsize=(w, h), dpi=250)
    ax.tick_params(length=0, labelsize=14, colors=INK); ax.grid(False)
    return fig, ax
def _save(fig, name):
    fig.tight_layout(); fig.savefig(OUT / name, facecolor="white"); plt.close(fig)
def hbar(labels, vals, name, w, h, xmax, fmt="{:.0f}%"):
    fig, ax = _fig(w, h)
    y = np.arange(len(labels))[::-1]
    ax.barh(y, vals, height=0.58, color=SEQ[4])
    for yi, v in zip(y, vals): ax.text(v + xmax * 0.015, yi, fmt.format(v), va="center", fontsize=15, color=INK, fontweight="bold")
    ax.set_yticks(y, labels, fontsize=15); ax.set_xlim(0, xmax); ax.set_xticks([])
    ax.spines["bottom"].set_visible(False)
    _save(fig, name)

hbar(list(data["providers"]), list(data["providers"].values()), "chart_providers.png", 7.2, 4.6, 70)
hbar(list(data["uni"]), list(data["uni"].values()), "chart_uni.png", 7.2, 4.6, 65)

fig, ax = _fig(7.4, 4.7)
lv = data["level"]; x = np.arange(len(lv))
ax.bar(x, list(lv.values()), width=0.58, color=SEQ[4])
for xi, v in zip(x, lv.values()): ax.text(xi, v + 0.8, str(v), ha="center", fontsize=15, color=INK, fontweight="bold")
ax.set_xticks(x, list(lv), fontsize=14); ax.set_yticks([]); ax.set_ylim(0, max(lv.values()) * 1.15)
_save(fig, "chart_level.png")

fig, ax = _fig(8.0, 4.7)
tv = list(data["time"].values()); x = np.arange(4)
ax.plot(x, tv, color=SEQ[4], lw=3.2, zorder=2)
ax.scatter(x, tv, s=[110, 110, 110, 260], color=[SEQ[4]] * 3 + [SEQ[6]], edgecolor="white", linewidth=2.2, zorder=3)
for xi, v in zip(x, tv): ax.text(xi, v + 3.2, f"{v:.0f}%", ha="center", fontsize=16, color=INK, fontweight="bold")
ax.set_xticks(x, ["Morning\n6–11 AM", "Afternoon\n12–3 PM", "Evening\n4–7 PM", "Night\nafter 8 PM"], fontsize=14)
ax.set_ylim(0, 60); ax.set_yticks([0, 20, 40, 60], ["0%", "20%", "40%", "60%"], color=INK2, fontsize=12)
ax.grid(axis="y", color=GRID, lw=1); ax.set_xlim(-0.35, 3.35)
_save(fig, "chart_time.png")

fig, ax = _fig(8.0, 4.75)
g = data["groups"]; names = ["Poor", "Average", "Good"]; x = np.arange(3); bw = 0.36
m = [g[k]["missed"] for k in names]; hi = [g[k]["high"] for k in names]
ax.bar(x - bw / 2 - 0.015, m, bw, color=SEQ[2], label="Missed a deadline because of connectivity")
ax.bar(x + bw / 2 + 0.015, hi, bw, color=SEQ[6], label="Academic performance highly affected")
for xs, vs in ((x - bw / 2 - 0.015, m), (x + bw / 2 + 0.015, hi)):
    for xi, v in zip(xs, vs): ax.text(xi, v + 2, f"{v:.0f}%", ha="center", fontsize=15, color=INK, fontweight="bold")
ax.set_xticks(x, [f"{k} connection\n({g[k]['n']} students)" for k in names], fontsize=14)
ax.set_ylim(0, 105); ax.set_yticks([0, 25, 50, 75, 100], ["0%", "25%", "50%", "75%", "100%"], color=INK2, fontsize=12)
ax.grid(axis="y", color=GRID, lw=1)
ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=2, frameon=False, fontsize=13, handlelength=1.1)
_save(fig, "chart_groups.png")
print("wrote slide chart images")

# ---- shrink images (few colours, so a 128-colour palette is visually lossless) to keep the upload small
from PIL import Image
for f in OUT.glob("*.png"):
    Image.open(f).convert("RGB").quantize(colors=128, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).save(f, optimize=True)
