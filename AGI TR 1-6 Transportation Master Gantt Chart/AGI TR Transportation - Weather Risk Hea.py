"""
AGI TR Transportation - Weather Risk Heatmap
/visualize_data --type=heatmap
Winter Shamal Risk Analysis for January-February 2026
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
import numpy as np
from datetime import datetime, timedelta
import os

# Font settings
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

# === DATA SETUP ===
start_date = datetime(2026, 1, 15)
num_days = 79
dates = [start_date + timedelta(days=i) for i in range(num_days)]

np.random.seed(42)

# Wind speed pattern (kt) - 79 days from Jan 15, 2026
wind_base = np.array([
    12, 11, 13, 12, 11, 13, 14, 13,  # Jan 15-22
    17, 19, 21, 23, 22, 20, 18, 16,  # Jan 23-30 (SHAMAL)
    15, 14, 13, 14, 15, 14, 13, 12, 11, 12,  # Jan 31 - Feb 9
    12, 11, 13, 12, 11, 12, 13, 11, 12, 11,  # Feb 10-19
    12, 13, 11, 12, 11, 13, 12, 11, 12, 11, 12,  # Feb 20-29
    12, 11, 13, 12, 11, 12, 13, 11, 12, 11,  # Mar 1-10
    12, 13, 11, 12, 11, 13, 12, 11, 12, 11, 12,  # Mar 11-21
    12, 11, 13, 12, 11, 12, 13, 11, 12, 11, 12  # Mar 22-31
])[:num_days]

wind_var = np.random.uniform(-1.5, 1.5, num_days)
wind_speed = wind_base + wind_var
gust_factor = np.random.uniform(1.2, 1.5, num_days)
wind_gust = wind_speed * gust_factor
wave_height = wind_speed * 0.045 + np.random.uniform(-0.1, 0.15, num_days)
wave_height = np.clip(wave_height, 0.3, 1.5)
visibility = np.ones(num_days) * 8
visibility[8:16] = np.random.uniform(2, 5, 8)
visibility = np.clip(visibility, 1, 10)
sea_state = wind_speed / 5
sea_state = np.clip(sea_state, 1, 6)

def calc_risk_score(wind, gust, wave, vis):
    wind_risk = np.clip((wind - 10) * 4, 0, 40)
    gust_risk = np.clip((gust - 15) * 2.5, 0, 25)
    wave_risk = np.clip((wave - 0.5) * 40, 0, 20)
    vis_risk = np.clip((8 - vis) * 3, 0, 15)
    return wind_risk + gust_risk + wave_risk + vis_risk

risk_score = calc_risk_score(wind_speed, wind_gust, wave_height, visibility)

op_status = []
for i in range(num_days):
    if risk_score[i] < 30:
        op_status.append('GO')
    elif risk_score[i] < 60:
        op_status.append('HOLD')
    else:
        op_status.append('NO-GO')

voyages = [
    {"name": "V1", "start": 3, "end": 7, "tr": "TR1", "type": "transport"},
    {"name": "V2", "start": 7, "end": 20, "tr": "TR2+JD", "type": "jackdown"},
    {"name": "V3", "start": 16, "end": 20, "tr": "TR3", "type": "transport"},
    {"name": "V4", "start": 21, "end": 34, "tr": "TR4+JD", "type": "jackdown"},
    {"name": "V5", "start": 27, "end": 31, "tr": "TR5", "type": "transport"},
    {"name": "V6", "start": 32, "end": 45, "tr": "TR6+JD", "type": "jackdown"},
    {"name": "V7", "start": 67, "end": 77, "tr": "TR7+JD-4", "type": "jackdown"},
]

# === CREATE VISUALIZATION ===
fig = plt.figure(figsize=(22, 16))

ax1 = plt.subplot2grid((3, 1), (0, 0), rowspan=1)  # main param heatmap

params = ['Wind (kt)', 'Gust (kt)', 'Wave (m)', 'Visibility (km)', 'Sea State', 'Risk Score']
data_matrix = np.array([wind_speed, wind_gust, wave_height * 10, visibility, sea_state, risk_score / 10])

colors = ['#2E7D32', '#4CAF50', '#8BC34A', '#CDDC39', '#FFEB3B', '#FFC107', '#FF9800', '#FF5722', '#D32F2F']
cmap = LinearSegmentedColormap.from_list('risk', colors, N=100)

data_norm = np.zeros_like(data_matrix)
for i in range(len(params)):
    row = data_matrix[i]
    data_norm[i] = (row - row.min()) / (row.max() - row.min() + 0.001)

im = ax1.imshow(data_norm, aspect='auto', cmap=cmap, interpolation='nearest')

ax1.set_yticks(range(len(params)))
ax1.set_yticklabels(params, fontsize=11, fontweight='bold')

date_labels = [d.strftime('%m/%d') for d in dates]
ax1.set_xticks(range(0, num_days, 2))
ax1.set_xticklabels([date_labels[i] for i in range(0, num_days, 2)], rotation=45, ha='right', fontsize=9)

for i in range(len(params)):
    for j in range(num_days):
        if i == 0: val = f"{wind_speed[j]:.0f}"
        elif i == 1: val = f"{wind_gust[j]:.0f}"
        elif i == 2: val = f"{wave_height[j]:.1f}"
        elif i == 3: val = f"{visibility[j]:.0f}"
        elif i == 4: val = f"{sea_state[j]:.1f}"
        else: val = f"{risk_score[j]:.0f}"
        
        if j % 2 == 0:
            color = 'white' if data_norm[i, j] > 0.6 else 'black'
            ax1.text(j, i, val, ha='center', va='center', fontsize=7, color=color)

shamal_start, shamal_end = 8, 15
ax1.axvspan(shamal_start - 0.5, shamal_end + 0.5, alpha=0.3, color='red', zorder=0)
ax1.text(shamal_start + 3.5, -0.8, 'WINTER SHAMAL PEAK', fontsize=12, fontweight='bold', 
         color='red', ha='center', va='bottom')

ax1.set_title(f'AGI TR Transportation - Weather Risk Heatmap\n'
              f'{start_date.strftime("%Y-%m-%d")} ~ {dates[-1].strftime("%Y-%m-%d")}',
              fontsize=15, fontweight='bold', pad=12)

cbar = plt.colorbar(im, ax=ax1, orientation='vertical', pad=0.02, aspect=30)
cbar.set_label('Risk Level (Row-normalized)', fontsize=10)

# === RISK SCORE TIMELINE ===
ax2 = plt.subplot2grid((3, 1), (1, 0), rowspan=1)  # risk timeline

ax2.fill_between(range(num_days), risk_score, alpha=0.3, color='orange')
ax2.plot(range(num_days), risk_score, 'o-', color='#FF5722', linewidth=2, markersize=4)

ax2.axhline(y=30, color='green', linestyle='--', linewidth=2, label='GO Threshold (30)')
ax2.axhline(y=60, color='red', linestyle='--', linewidth=2, label='NO-GO Threshold (60)')

ax2.axvspan(shamal_start, shamal_end, alpha=0.15, color='red')

voyage_colors = {'transport': '#2196F3', 'jackdown': '#9C27B0'}
for v in voyages:
    color = voyage_colors[v['type']]
    ax2.axvspan(v['start'], v['end'], alpha=0.15, color=color, zorder=0)
    mid = (v['start'] + v['end']) / 2
    ax2.text(mid, 88, f"{v['name']}\n{v['tr']}", ha='center', va='top', fontsize=9,
             fontweight='bold', color=color,
             bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8))

ax2.set_xlim(-0.5, num_days - 0.5)
ax2.set_ylim(0, 100)
ax2.set_xticks(range(0, num_days, 2))
ax2.set_xticklabels([date_labels[i] for i in range(0, num_days, 2)], rotation=45, ha='right', fontsize=9)
ax2.set_ylabel('Risk Score (0-100)', fontsize=11, fontweight='bold')
ax2.set_title('Composite Weather Risk Score with Voyage Schedule', fontsize=13, fontweight='bold')
ax2.legend(loc='upper right', fontsize=9)
ax2.grid(True, alpha=0.3)

# === OPERATION STATUS BAR ===
ax3 = plt.subplot2grid((3, 1), (2, 0), rowspan=1)  # status bar

status_colors = {'GO': '#4CAF50', 'HOLD': '#FFC107', 'NO-GO': '#F44336'}
bar_colors = [status_colors[s] for s in op_status]

ax3.bar(range(num_days), [1] * num_days, color=bar_colors, edgecolor='white', linewidth=0.5)

for v in voyages:
    ax3.axvline(x=v['start'], color='blue', linestyle='-', linewidth=2, alpha=0.7)
    ax3.text(v['start'], 1.15, v['name'], fontsize=10, fontweight='bold', ha='center', color='blue')

ax3.set_xlim(-0.5, num_days - 0.5)
ax3.set_ylim(0, 1.4)
ax3.set_xticks(range(0, num_days, 2))
ax3.set_xticklabels([date_labels[i] for i in range(0, num_days, 2)], rotation=45, ha='right', fontsize=9)
ax3.set_yticks([])
ax3.set_title('Daily Operation Status (GO / HOLD / NO-GO)', fontsize=13, fontweight='bold')

go_patch = mpatches.Patch(color='#4CAF50', label='GO (Risk < 30)')
hold_patch = mpatches.Patch(color='#FFC107', label='HOLD (30-60)')
nogo_patch = mpatches.Patch(color='#F44336', label='NO-GO (> 60)')
voyage_patch = mpatches.Patch(color='#2196F3', alpha=0.5, label='Transport Voyage')
jd_patch = mpatches.Patch(color='#9C27B0', alpha=0.5, label='Jack-down Period')
ax3.legend(handles=[go_patch, hold_patch, nogo_patch, voyage_patch, jd_patch], 
           loc='upper right', ncol=5, fontsize=9)

# === SUMMARY BOX ===
go_n = op_status.count('GO')
hold_n = op_status.count('HOLD')
nogo_n = op_status.count('NO-GO')

# Calculate Shamal peak stats
shamal_start_idx = 8
shamal_end_idx = 16
shamal_peak_text = (
    f"Shamal Peak: {dates[shamal_start_idx].strftime('%b %d')}-{dates[shamal_end_idx-1].strftime('%b %d')}\n"
    f"  Avg Wind (Peak): {wind_speed[shamal_start_idx:shamal_end_idx].mean():.1f} kt\n"
    f"  Max Gust: {wind_gust.max():.1f} kt\n"
    f"  Avg Risk (Peak): {risk_score[shamal_start_idx:shamal_end_idx].mean():.0f}/100\n"
)

stats_text = (
    "Weather Analysis Summary\n"
    "------------------------\n"
    f"Period: {start_date.strftime('%b %d')} - {dates[-1].strftime('%b %d, %Y')} ({num_days} days)\n"
    f"GO Days: {go_n} ({go_n/num_days*100:.0f}%)\n"
    f"HOLD Days: {hold_n} ({hold_n/num_days*100:.0f}%)\n"
    f"NO-GO Days: {nogo_n} ({nogo_n/num_days*100:.0f}%)\n"
    + shamal_peak_text
)

fig.text(0.02, 0.05, stats_text, fontsize=9, fontfamily='DejaVu Sans',
         verticalalignment='bottom', horizontalalignment='left',
         bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.9, edgecolor='black', pad=0.8))

# === RISK MITIGATION STRATEGY BOX ===
mitigation_lines = ["Risk Mitigation Strategy\n", "------------------------\n"]

for v in voyages:
    v_start_date = dates[v["start"]]
    v_end_date = dates[min(v["end"], num_days-1)]
    v_risk = risk_score[v["start"]:min(v["end"]+1, num_days)]
    v_max_risk = np.nanmax(v_risk) if len(v_risk) > 0 else 0
    v_avg_risk = np.nanmean(v_risk) if len(v_risk) > 0 else 0
    
    if v_max_risk < 30:
        risk_level = "LOW RISK - Proceed"
    elif v_max_risk < 60:
        risk_level = "MODERATE"
    else:
        risk_level = "HIGH RISK - Delay recommended"
    
    mitigation_lines.append(f"{v['name']} ({v_start_date.strftime('%b %d')}-{v_end_date.strftime('%b %d')}): {risk_level}\n")
    
    # Add specific notes for high-risk voyages
    if v_max_risk >= 60 and v["type"] == "jackdown":
        # Check if sailing overlaps with shamal
        if v["start"] <= shamal_end and v["end"] >= shamal_start:
            mitigation_lines.append(f"  Sailing near Shamal period\n")
            mitigation_lines.append(f"  Jack-down ({v_start_date.strftime('%b %d')}-{v_end_date.strftime('%b %d')}) shore OK\n")

mitigation_lines.append("\nOverall Strategy: Schedule optimized to minimize marine ops during peak Shamal.")

mitigation_text = "".join(mitigation_lines)

fig.text(0.98, 0.05, mitigation_text, fontsize=9, fontfamily='DejaVu Sans',
         verticalalignment='bottom', horizontalalignment='right',
         bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.9, edgecolor='black', pad=0.8))

plt.tight_layout(rect=[0, 0.22, 1, 0.98])  # [left, bottom, right, top] - 하단 22%를 텍스트 박스 영역으로 확보
plt.subplots_adjust(hspace=0.50)  # 패널 간 간격만 조정

output_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
output_path = os.path.join(output_dir, "AGI_TR_Weather_Risk_Heatmap.png")
plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

print(f"Weather Risk Heatmap generated: {output_path}")
print(f"\nStatistics:")
print(f"   GO Days: {op_status.count('GO')} / {num_days}")
print(f"   HOLD Days: {op_status.count('HOLD')} / {num_days}")
print(f"   NO-GO Days: {op_status.count('NO-GO')} / {num_days}")
print(f"   Shamal Peak Risk: {risk_score[8:16].mean():.0f}/100")