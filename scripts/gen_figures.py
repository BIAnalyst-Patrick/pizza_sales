#!/usr/bin/env python3
"""Report figures for the Plato's Pizza analysis. Palette: dataviz reference (validated)."""
import pandas as pd, numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import os

SRC = "data/pizza_sales.csv"
FIG = "report/figures"
os.makedirs(FIG, exist_ok=True)

ACCENT = "#EB6834"; INK = "#0B0B0B"; INK2 = "#52514E"; MUTED = "#898781"
GRID = "#E1E0D9"; SURFACE = "#FCFCFB"
CAT = {"Chicken": "#2A78D6", "Classic": "#1BAF7A", "Supreme": "#EDA100", "Veggie": "#008300"}

plt.rcParams.update({
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE,
    "axes.edgecolor": "#C3C2B7", "axes.linewidth": 0.8,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6,
    "font.family": "DejaVu Sans", "text.color": INK,
    "axes.labelcolor": INK2, "xtick.color": MUTED, "ytick.color": MUTED,
    "axes.titlesize": 13, "axes.titleweight": "bold", "axes.titlecolor": INK,
    "axes.titlelocation": "left", "font.size": 9.5,
    "axes.spines.top": False, "axes.spines.right": False,
    "savefig.dpi": 160, "savefig.bbox": "tight", "savefig.facecolor": SURFACE,
})

df = pd.read_csv(SRC)
df["order_date"] = pd.to_datetime(df["order_date"], format="%d-%m-%Y")
df["hour"] = pd.to_datetime(df["order_time"], format="%H:%M:%S").dt.hour
df["dow"] = df["order_date"].dt.dayofweek
df["week"] = df["order_date"].dt.isocalendar().week.astype(int)

def short(n): return n.replace("The ", "").replace(" Pizza", "")

# ---- Fig 1: the plateau ----------------------------------------------------
weekly = df.groupby("week")["total_price"].sum()
opend = df.groupby("week")["order_date"].nunique()
wpd = weekly / opend
avg = df["total_price"].sum() / df["order_date"].nunique()
fig, ax = plt.subplots(figsize=(9, 3.6))
ax.plot(wpd.index, wpd.values, color=ACCENT, lw=2, solid_capstyle="round")
ax.axhline(avg, color=MUTED, lw=1.4, ls=(0, (4, 3)))
ax.annotate(f"year average ${avg:,.0f}/open day", xy=(53, avg), xytext=(38.5, avg + 320),
            fontsize=9, color=INK2)
q = df.groupby(df["order_date"].dt.quarter).apply(lambda g: g["total_price"].sum() / g["order_date"].nunique(), include_groups=False)
for qi, (wk, val) in enumerate(zip([7, 20, 33, 46], q.values)):
    ax.annotate(f"Q{qi+1}: ${val:,.0f}", xy=(wk, 1350), fontsize=8.5, color=MUTED, ha="center")
ax.set_title("Revenue has flatlined at ≈ $2,285 per open day", pad=14)
ax.set_xlabel("ISO week, 2015"); ax.set_ylabel("Revenue per open day ($)")
ax.set_ylim(1200, 3200); ax.set_xlim(1, 53)
fig.savefig(f"{FIG}/fig1_plateau.png"); plt.close(fig)

# ---- Fig 2: menu engineering matrix -----------------------------------------
pz = df.groupby(["pizza_name", "pizza_category"]).agg(units=("quantity", "sum"), rev=("total_price", "sum")).reset_index()
pz["price"] = pz.rev / pz.units
mu, mp = pz.units.median(), pz.price.median()
fig, ax = plt.subplots(figsize=(9, 6.4))
for cat_name, g in pz.groupby("pizza_category"):
    ax.scatter(g.units, g.price, s=70, color=CAT[cat_name], label=cat_name,
               edgecolor=SURFACE, linewidth=1.2, zorder=3)
ax.axvline(mu, color=MUTED, lw=1, ls=(0, (4, 3))); ax.axhline(mp, color=MUTED, lw=1, ls=(0, (4, 3)))
xmax, ymin, ymax = pz.units.max() * 1.08, pz.price.min() - 1, pz.price.max() + 1
qlab = dict(fontsize=11, fontweight="bold", color="#B8B6AF", zorder=1)
ax.text(mu * 1.72, ymax - 0.5, "STARS — protect & feature", ha="center", **qlab)
ax.text(mu * 1.72, ymin + 0.4, "PLOWHORSES — upsell / reprice", ha="center", **qlab)
ax.text(mu * 0.42, ymax - 0.5, "PUZZLES — promote", ha="center", **qlab)
ax.text(mu * 0.42, ymin + 0.4, "DOGS — fix or cut", ha="center", **qlab)
for name, dx, dy in [("The Brie Carre Pizza", 8, 4), ("The Thai Chicken Pizza", 0, 9),
                     ("The Pepperoni Pizza", 0, -13), ("The Greek Pizza", 8, 5),
                     ("The Big Meat Pizza", -20, 8), ("The Hawaiian Pizza", 0, 9),
                     ("The Green Garden Pizza", 8, -11)]:
    r = pz[pz.pizza_name == name].iloc[0]
    ax.annotate(short(name), (r.units, r.price), xytext=(dx, dy), textcoords="offset points",
                fontsize=8.5, color=INK2)
ax.set_title("The menu has four personalities — manage each differently", pad=14)
ax.set_xlabel("Units sold, 2015"); ax.set_ylabel("Average realised price per pizza ($)")
ax.set_xlim(0, xmax); ax.set_ylim(ymin, ymax)
leg = ax.legend(ncol=4, loc="lower left", bbox_to_anchor=(0, 1.02), frameon=False, fontsize=9, title=None)
ax.set_title("The menu has four personalities — manage each differently", pad=32)
fig.savefig(f"{FIG}/fig2_menu_matrix.png"); plt.close(fig)

# ---- Fig 3: demand heatmap ---------------------------------------------------
days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
hm = df[(df.hour >= 11) & (df.hour <= 23)].pivot_table(
    index="hour", columns="dow", values="total_price", aggfunc="sum", fill_value=0)
fig, ax = plt.subplots(figsize=(7.6, 5.2))
im = ax.imshow(hm.values, cmap=matplotlib.colors.LinearSegmentedColormap.from_list(
    "seq", ["#FFFFFF", "#9EC5F4", "#256ABF", "#104281"]), aspect="auto")
ax.set_xticks(range(7), days); ax.set_yticks(range(len(hm.index)), [f"{h:02d}:00" for h in hm.index])
ax.grid(False)
for i in range(hm.shape[0]):
    for j in range(hm.shape[1]):
        v = hm.values[i, j]
        ax.text(j, i, f"{v/1000:.1f}" if v >= 100 else "·", ha="center", va="center",
                fontsize=7.5, color="#FFFFFF" if v > hm.values.max() * 0.55 else INK2)
ax.set_title("When the money arrives — revenue by hour × weekday ($K)", pad=14)
cb = fig.colorbar(im, ax=ax, shrink=0.7, format=lambda x, _: f"${x/1000:.0f}K")
cb.outline.set_visible(False)
fig.savefig(f"{FIG}/fig3_heatmap.png"); plt.close(fig)

# ---- Fig 4: sorted revenue bars ----------------------------------------------
par = pz.sort_values("rev", ascending=False).reset_index(drop=True)
par["cum"] = par.rev.cumsum() / par.rev.sum()
top_n = int((par.cum <= 0.50).sum()) + 1
colors = [ACCENT if i < top_n else ("#E1E0D9" if i >= len(par) - 10 else "#C3C2B7") for i in range(len(par))]
fig, ax = plt.subplots(figsize=(9, 7.2))
ax.barh([short(n) for n in par.pizza_name][::-1], par.rev[::-1] / 1000, color=colors[::-1], height=0.72)
ax.set_title(f"Half the revenue comes from just {top_n} of 32 pizzas — the bottom 10 add 19%", pad=14)
ax.set_xlabel("Revenue 2015 ($K)")
ax.tick_params(axis="y", labelsize=8.5)
ax.axvline(0, color="#C3C2B7", lw=0.8)
for lbl, txt in [(31.6, f"top {top_n} = 50% of revenue"), (5.5, "bottom 10 = 19%")]:
    pass
ax.annotate(f"top {top_n} pizzas = 50% of revenue", xy=(29, 16), fontsize=9, color=ACCENT, fontweight="bold")
ax.annotate("bottom 10 = 19% — review, fix or cut\n(XXL Greek: 28 units all year)", xy=(14, 4), fontsize=9, color=INK2)
fig.savefig(f"{FIG}/fig4_pareto.png"); plt.close(fig)

# ---- Fig 5: opportunity waterfall ---------------------------------------------
lad = df.pivot_table(index="pizza_name", columns="pizza_size", values="unit_price", aggfunc="mean")
su = df.groupby("pizza_size")["quantity"].sum()
w_upsell = 0.20 * su["S"] * (lad["M"] - lad["S"]).mean() + 0.20 * su["M"] * (lad["L"] - lad["M"]).mean()
osz = df.groupby("order_id")["quantity"].sum()
w_attach = (osz == 1).sum() * 0.10 * df.loc[df.pizza_size == "S", "unit_price"].mean()
w_happy = 0.30 * df.loc[df.hour == 15, "total_price"].sum()
fri = df[df.dow == 4].groupby("order_date")["total_price"].sum().mean()
sun = df[df.dow == 6].groupby("order_date")["total_price"].sum().mean()
w_weekend = 0.25 * (fri - sun) * df.loc[df.dow == 6, "order_date"].nunique()
puz = pz[(pz.units < mu) & (pz.price >= mp)]
w_puzzle = 0.15 * puz.rev.sum()
items = [("Size upsell\n(1-in-5 upgrade)", w_upsell), ("2nd-pizza attach\n(10% of solo orders)", w_attach),
         ("3–4pm happy hour\n(+30% on trough)", w_happy), ("Close Sunday gap\n(25% of Fri–Sun gap)", w_weekend),
         ("Promote Puzzles\n(+15% volume)", w_puzzle)]
total = sum(v for _, v in items)
fig, ax = plt.subplots(figsize=(9, 4.6))
cum = 0
for i, (name, v) in enumerate(items):
    ax.bar(i, v / 1000, bottom=cum / 1000, color=ACCENT, width=0.62)
    ax.text(i, (cum + v) / 1000 + 2, f"+${v/1000:,.1f}K", ha="center", fontsize=9, color=INK, fontweight="bold")
    cum += v
ax.bar(len(items), total / 1000, color="#1A1A19", width=0.62)
ax.text(len(items), total / 1000 + 2, f"+${total/1000:,.0f}K\n(+11%)", ha="center", fontsize=9.5, color=INK, fontweight="bold")
ax.set_xticks(range(len(items) + 1), [n for n, _ in items] + ["Total\nopportunity"], fontsize=8.5)
ax.set_title("The path to +$88K — five conservative levers, no new customers", pad=14)
ax.set_ylabel("Annualised revenue gain ($K)")
ax.set_ylim(0, 105)
fig.savefig(f"{FIG}/fig5_waterfall.png"); plt.close(fig)

print("figures saved:", sorted(os.listdir(FIG)))
print(f"levers: upsell={w_upsell:,.0f} attach={w_attach:,.0f} happy={w_happy:,.0f} weekend={w_weekend:,.0f} puzzle={w_puzzle:,.0f} total={total:,.0f}")
print(f"top_n for 50%: {top_n}")
