import numpy as np
import pandas as pd
import os
from matplotlib import pyplot as plt
import matplotlib.font_manager as fm
from openpyxl.reader.excel import load_workbook
import os

#Add basepath
basepath = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

# --- Configurable options ---
metric = "costs"       # Choose: "costs" or "emissions"
scale_type = "per_tonne"   # Choose: "total" or "per_tonne"
saveas = 'pdf'  # Options: "no", "svg", "pdf", "both"
filename = 'TC_baseline_' + metric + '_' + scale_type + '_noCO2'
delayed = 0
sensitivity = 1
stacked = 0

node = "Chemelot"
if node == "Chemelot":
    if sensitivity:
        file_path = os.path.join(basepath, "Plotting", "result_data_long_noCO2.xlsx")
    else:
        file_path = os.path.join(basepath, "Plotting", "result_data_long.xlsx")
else:
    file_path = os.path.join(basepath, "Plotting", "result_data_long_Zeeland.xlsx")

# Set font to Open Sans
# font_path = 'C:/Windows/Fonts/OpenSans-Regular.ttf'  # Make sure this path is correct
# if os.path.exists(font_path):
#     open_sans = fm.FontProperties(fname=font_path)
#     plt.rcParams['font.family'] = open_sans.get_name()
# else:
#     print("Open Sans font not found, using default.")
#     plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.family'] = 'serif'

# --- Load and prepare data ---
df = pd.read_excel(file_path, sheet_name="Sheet1", header=None)
top_header = df.iloc[0].ffill()
sub_header = df.iloc[2].ffill()
df.columns = pd.MultiIndex.from_arrays([top_header, sub_header])
top_header = top_header.copy()
top_header.iloc[0] = ""
sub_header.iloc[0] = "Result type"
df = df.iloc[4:].reset_index(drop=True)

keywords = ["costs_obj_interval", "sunk_costs", "costs_tot_interval", "costs_tot_cumulative", "emissions_net"]
if node == "Chemelot":
    total_product = 3089300
    total_product_cum = total_product * 30
else:
    total_product = 4461016
    total_product_cum = total_product * 30
first_column_name = df.columns[0]
filtered_df = df[df[first_column_name].astype(str).str.startswith(tuple(keywords))]

def get_val(df, label, year, cost_type):
    try:
        return float(df.loc[df[first_column_name] == cost_type, (label, year)].values[0])
    except:
        return 0


# --- Extract values ---
greenfield_limit = [get_val(filtered_df, 'EmissionLimit Greenfield', y, "costs_tot_cumulative") / 1e6 for y in ['2030', '2040', '2050']]
brownfield_limit = [get_val(filtered_df, 'EmissionLimit Brownfield', y, "costs_tot_interval") * 10 /1e6  for y in ['2030', '2040', '2050']]
brownfield_limit_sum = sum(brownfield_limit)

if node == "Chemelot" and not sensitivity:
    greenfield_scope = [get_val(filtered_df, 'EmissionScope Greenfield', y, "costs_tot_cumulative") / 1e6 for y in ['2030', '2040', '2050']]
    brownfield_scope = [get_val(filtered_df, 'EmissionScope Brownfield', y, "costs_tot_interval") * 10 / 1e6 for y in ['2030', '2040', '2050']]
    brownfield_scope_sum = sum(brownfield_scope)

total_em_limit = [get_val(filtered_df, 'EmissionLimit Greenfield', y, "emissions_net") * 30 / 1e6 for y in ['2030', '2040', '2050']]
total_em_limit_bf = [get_val(filtered_df, 'EmissionLimit Brownfield', y, "emissions_net") * 10 / 1e6 for y in ['2030', '2040', '2050']]
if node == "Chemelot" and not sensitivity:
    total_em_scope = [get_val(filtered_df, 'EmissionScope Greenfield', y, "emissions_net") * 30 / 1e6 for y in ['2030', '2040', '2050']]
    total_em_scope_bf = [get_val(filtered_df, 'EmissionScope Brownfield', y, "emissions_net") * 10 / 1e6 for y in ['2030', '2040', '2050']]

# --- Data Selection ---
if metric == "costs":
    ylabel_base = "Cluster costs"
    unit = "M€" if scale_type == "total" else "€/tonne product"
    gf_limit_vals = greenfield_limit
    bf_limit_vals = brownfield_limit
    if node == "Chemelot" and not sensitivity:
        gf_scope_vals = greenfield_scope
        bf_scope_vals = brownfield_scope
    if scale_type == "per_tonne":
        #recalculate for brownfield
        # bf_limit_total = sum(total_em_limit_bf)
        # bf_limit_avg = sum(total_em_limit_bf) / len(total_em_limit_bf)
        # bf_limit_fractions = [(v / bf_limit_total if bf_limit_total > 0 else 0) for v in bf_limit_vals]
        # bf_scope_total = sum(total_em_scope_bf)
        # bf_scope_avg = sum(total_em_scope_bf) / len(total_em_scope_bf)
        # bf_scope_fractions = [(v / bf_scope_total if bf_scope_total > 0 else 0) for v in bf_scope_vals]

        gf_limit_vals = [val * 1e6 / total_product_cum for val in gf_limit_vals]
        bf_limit_vals = [val * 1e6 / (total_product * 10) for val in bf_limit_vals]
        if node == "Chemelot" and not sensitivity:
            gf_scope_vals = [val * 1e6 / total_product_cum for val in gf_scope_vals]
            bf_scope_vals = [val * 1e6 / (total_product * 10) for val in bf_scope_vals]
else:
    ylabel_base = "Cluster emissions"
    unit = "MtCO$_2$" if scale_type == "total" else "tCO$_2$/tonne product"
    gf_limit_vals = total_em_limit
    bf_limit_vals = total_em_limit_bf
    if node == "Chemelot":
        gf_scope_vals = total_em_scope
        bf_scope_vals = total_em_scope_bf
    if scale_type == "per_tonne":
        #recalculate for brownfield
        # bf_limit_total = sum(total_em_limit_bf)
        # bf_limit_avg = sum(total_em_limit_bf) / len(total_em_limit_bf)
        # bf_limit_fractions = [(v / bf_limit_total if bf_limit_total > 0 else 0) for v in bf_limit_vals]
        # bf_scope_total = sum(total_em_scope_bf)
        # bf_scope_avg = sum(total_em_scope_bf) / len(total_em_scope_bf)
        # bf_scope_fractions = [(v / bf_scope_total if bf_scope_total > 0 else 0) for v in bf_scope_vals]

        gf_limit_vals = [val * 1e6 / total_product_cum for val in gf_limit_vals]
        bf_limit_vals = [val * 1e6 / (total_product * 10) for val in bf_limit_vals]
        if node == "Chemelot":
            gf_scope_vals = [val * 1e6 / total_product_cum for val in gf_scope_vals]
            bf_scope_vals = [val * 1e6 / (total_product * 10) for val in bf_scope_vals]

if stacked:
    # --- Plotting ---
    if node == "Chemelot" and not sensitivity:
        fig, ax = plt.subplots(figsize=(10, 3))
        x = np.arange(8)
    else:
        fig, ax = plt.subplots(figsize=(5, 3))
        x = np.arange(4)

    colors = ["#7F9183", "#7F9183", "#7F9183", "#7F9183"]
    brownfield_color = ["#765B56", "#C3B0AC", "#DDD5D0"]
    x_labels = ['Greenfield\n2030', 'Greenfield\n2040', 'Greenfield\n2050', 'Brownfield',
                'Greenfield\n2030', 'Greenfield\n2040', 'Greenfield\n2050', 'Brownfield']

    # EmissionLimit / Cost bars
    ax.bar(x[:3], gf_limit_vals, color=colors, label='Greenfield')
    ax.bar(x[3], bf_limit_vals[0], color=brownfield_color[0])
    ax.bar(x[3], bf_limit_vals[1], bottom=bf_limit_vals[0], color=brownfield_color[1])
    ax.bar(x[3], bf_limit_vals[2], bottom=bf_limit_vals[0] + bf_limit_vals[1], color=brownfield_color[2])

    # EmissionScope / Cost bars
    if node == "Chemelot" and not sensitivity:
        ax.bar(x[4:7], gf_scope_vals, color=colors)
        ax.bar(x[7], bf_scope_vals[0], color=brownfield_color[0])
        ax.bar(x[7], bf_scope_vals[1], bottom=bf_scope_vals[0], color=brownfield_color[1])
        ax.bar(x[7], bf_scope_vals[2], bottom=bf_scope_vals[0] + bf_scope_vals[1], color=brownfield_color[2])

        # Dashed line between sections
        ax.axvline(x=3.5, color='black', linestyle='dashed')

    # Labels
    ax.set_xticks(x)
    ax.set_xticklabels(x_labels)
    ax.set_ylabel(f"{ylabel_base} [{unit}]")
    ax.text(1.5, ax.get_ylim()[1] * 1.05, "Scope 1, 2 & 3", ha='center', fontsize=10) #change for text position
    if node == "Chemelot" and not sensitivity:
        ax.text(5.5, ax.get_ylim()[1] * 1.05, "Scope 1 & 2", ha='center', fontsize=10)

    # Legend
    custom_legend = [
        plt.Rectangle((0, 0), 1, 1, color=colors[0], label='Greenfield'),
        plt.Rectangle((0, 0), 1, 1, color=brownfield_color[0], label='Brownfield (2030)'),
        plt.Rectangle((0, 0), 1, 1, color=brownfield_color[1], label='Brownfield (2040)'),
        plt.Rectangle((0, 0), 1, 1, color=brownfield_color[2], label='Brownfield (2050)'),
    ]
    ax.legend(handles=custom_legend, loc='upper left')

    # #axes limit
    # if scale_type == 'per_tonne':
    #     if metric == 'costs':
    #         ax.set_ylim(0, 1500)
    #     else:
    #         ax.set_ylim(0, 0.28)
    # else:
    #     if metric == 'costs':
    #         ax.set_ylim(0, 1500)
    #     else:
    #         ax.set_ylim(0, 0.3)


else:
    # --- Plotting ---
    if node == "Chemelot" and not sensitivity:
        fig, ax = plt.subplots(figsize=(10, 3.5))
    else:
        fig, ax = plt.subplots(figsize=(5, 3.5))

    years = ['2030', '2040', '2050']
    n_years = len(years)

    # X positions
    group_width = 1  # space between year groups
    bar_width = 0.35
    x_greenfield = np.arange(n_years) * group_width - bar_width / 2
    x_brownfield = np.arange(n_years) * group_width + bar_width / 2

    # Colors
    # greenfield_color = "#7F9183"
    greenfield_color = "#87bd73"
    brownfield_color = "#4d3c30"

    # --- EmissionLimit bars ---
    ax.bar(x_greenfield[0], gf_limit_vals[0], width=bar_width, color=greenfield_color, label='Greenfield')
    ax.bar(x_brownfield[0], bf_limit_vals[0], width=bar_width, color=brownfield_color)

    ax.bar(x_greenfield[1], gf_limit_vals[1], width=bar_width, color=greenfield_color)
    ax.bar(x_brownfield[1], bf_limit_vals[1], width=bar_width, color=brownfield_color)

    ax.bar(x_greenfield[2], gf_limit_vals[2], width=bar_width, color=greenfield_color)
    ax.bar(x_brownfield[2], bf_limit_vals[2], width=bar_width, color=brownfield_color)

    # --- EmissionScope bars ---
    if node == "Chemelot" and not sensitivity:
        offset = group_width * (n_years) + 1.5 # space between Limit and Scope

        ax.bar(x_greenfield[0] + offset, gf_scope_vals[0], width=bar_width, color=greenfield_color)
        ax.bar(x_brownfield[0] + offset, bf_scope_vals[0], width=bar_width, color=brownfield_color)

        ax.bar(x_greenfield[1] + offset, gf_scope_vals[1], width=bar_width, color=greenfield_color)
        ax.bar(x_brownfield[1] + offset, bf_scope_vals[1], width=bar_width, color=brownfield_color)

        ax.bar(x_greenfield[2] + offset, gf_scope_vals[2], width=bar_width, color=greenfield_color)
        ax.bar(x_brownfield[2] + offset, bf_scope_vals[2], width=bar_width, color=brownfield_color)

    # --- Decorations ---
    if node == "Chemelot" and not sensitivity:
        xticks_positions = list((x_greenfield + x_brownfield) / 2) + list(
            ((x_greenfield + offset) + (x_brownfield + offset)) / 2)
        xticks_labels = years + years
    else:
        xticks_positions = list((x_greenfield + x_brownfield) / 2)
        xticks_labels = years

    ax.set_xticks(xticks_positions)
    ax.set_xticklabels(xticks_labels)

    # Dashed line between Limit and Scope
    if node == "Chemelot" and not sensitivity:
        ax.axvline(x=(x_brownfield[2] + x_greenfield[0] + offset) / 2, color='black', linestyle='dashed')

    # Horizontal price line (only for costs)
    if metric == "costs":
        if node == "Chemelot" and not sensitivity:
            price_line = ax.axhline(y=880, color='#4c4d4c', linestyle='--', linewidth=1)
        else:
            price_line = ax.axhline(y=868, color='#4c4d4c', linestyle='--', linewidth=1)

    # Y-axis label
    ax.set_ylabel(f"{ylabel_base} [{unit}]")
    # ax.set_ylim(0, 1900)

    # Section labels
    if node == "Chemelot" and not sensitivity:
        ax.text((x_brownfield[2] + x_greenfield[0]) / 2, ax.get_ylim()[1] * 1.05, "Scope 1, 2 & 3", ha='center',
                fontsize=10)
        ax.text((x_brownfield[2] + x_greenfield[0] + 2 * offset) / 2, ax.get_ylim()[1] * 1.05, "Scope 1 & 2", ha='center',
            fontsize=10)

    # --- Custom legend ---
    custom_legend = [
        plt.Rectangle((0, 0), 1, 1, color=greenfield_color, label='Greenfield'),
        plt.Rectangle((0, 0), 1, 1, color=brownfield_color, label='Brownfield'),
    ]

    if metric == "costs":
        custom_legend.append(
            plt.Line2D([0], [0], color='#4c4d4c', linestyle='--', linewidth=1, label='Weighted average\nproduct '
                                                                                  'price 2024')
        )
    # if node == "Chemelot" and not sensitivity:
    #     ax.legend(handles=custom_legend, loc='upper center')
    # else:
    #     ax.legend(handles=custom_legend, loc='upper right')

#Print costs and difference
print(bf_limit_vals)
print(gf_limit_vals)
if node == "Chemelot" and not sensitivity:
    print(gf_scope_vals)
    print(bf_scope_vals)

for i in [0, 1, 2]:
    ans = (bf_limit_vals[i] - gf_limit_vals[i]) / gf_limit_vals[i]
    print(ans)
if node == "Chemelot" and not sensitivity:
    for i in [0, 1, 2]:
        ans = (bf_scope_vals[i] - gf_scope_vals[i]) / gf_scope_vals[i]
        print(ans)



# --- Optional saving ---
if node == "Zeeland":
    filename = filename + "_" + node
if saveas in ['svg', 'pdf', 'both']:
    savebasepath = os.path.join(DATAPATH, "Plotting", "MY_Plots")
    if saveas in ['pdf', 'both']:
        plt.savefig(os.path.join(savebasepath, f"{filename}.pdf"), format='pdf')
    if saveas in ['svg', 'both']:
        plt.savefig(os.path.join(savebasepath, f"{filename}.svg"), format='svg')

# plt.tight_layout(h_pad=1.5)
plt.tight_layout()
plt.show()
