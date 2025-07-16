import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
from matplotlib.lines import Line2D
from cmcrameri import cm
import os

#Add basepath
basepath = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

# Define the data path
datapath = os.path.join(basepath, "Input_data", "250303_MY_Data_Chemelot")
savepath = os.path.join(basepath, "Plotting", "MY_Plots")
el_load_path = Path(datapath) / 'import_data' / 'Electricity_data_MY.xlsx'
save = 1

# Years to plot and their colors
nodes = ['Chemelot', 'Zeeland']
years = ['2030', '2040', '2050']
labels = {'2030': '2030', '2040': '2040', '2050': '2050'}
# labels = {'2030': 'Short-term', '2040': 'Mid-term', '2050': 'Long-term'}
colors_chemelot = cm.lipari(np.linspace(0.4, 0.8, len(years)))
colors_zeeland = cm.devon(np.linspace(0, 0.8, len(years)))
colors_emission = cm.grayC(np.linspace(0, 0.8, len(years)))

# Dictionary to store data
data = {node: {} for node in nodes}

# Load data for each year
for year in years:
    el_importdata = pd.read_excel(el_load_path, sheet_name=year, header=0, nrows=8760)
    for node in nodes:
        if node == 'Chemelot':
            data[node][year] = {
                'price': el_importdata.iloc[:, 0],
                'emission': el_importdata.iloc[:, 2] * 1000
            }
        elif node == 'Zeeland':
            data[node][year] = {
                'price': el_importdata.iloc[:, 1],
                'emission': el_importdata.iloc[:, 2] * 1000
            }

# Configure Matplotlib to use LaTeX for text rendering and set font
plt.rcParams.update({
    "text.usetex": True,
    "font.family": "serif",
    "font.serif": ["Computer Modern Roman"],
    "axes.labelsize": 10,
    "font.size": 10,
    "legend.fontsize": 10,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
})

# === FIGURE 1: HOURLY PROFILES ===
fig1, axes1 = plt.subplots(3, 1, figsize=(7, 7), sharex=True)
fig1.subplots_adjust(hspace=0.4, top=0.95, bottom=0.1, left=0.1, right=0.98)

# Chemelot Price
ax1 = axes1[0]
for i, year in enumerate(years):
    ax1.plot(data['Chemelot'][year]['price'], color=colors_chemelot[i],
             label=f"{labels[year]} ({data['Chemelot'][year]['price'].mean():.1f} EUR/MWh)", linewidth=0.5)
ax1.set_title('Electricity Prices: Chemelot')
ax1.set_ylabel('Electricity Price [EUR/MWh]')
ax1.set_xlim(0, 8759)
ax1.set_ylim(0, 200)
ax1.legend(loc='upper right')

# Zeeland Price
ax2 = axes1[1]
for i, year in enumerate(years):
    ax2.plot(data['Zeeland'][year]['price'], color=colors_zeeland[i],
             label=f"{labels[year]} ({data['Zeeland'][year]['price'].mean():.1f} EUR/MWh)", linewidth=0.5)
ax2.set_title('Electricity Prices: Zeeland')
ax2.set_ylabel('Electricity Price [EUR/MWh]')
ax2.set_xlim(0, 8759)
ax2.set_ylim(0, 200)
ax2.legend(loc='upper right')

# Emission Rate (assume same for both nodes)
ax3 = axes1[2]
for i, year in enumerate(years):
    ax3.plot(data['Chemelot'][year]['emission'], color=colors_emission[i],
             label=f"{labels[year]} ({data['Chemelot'][year]['emission'].mean():.0f} kg CO$_2$/MWh)", linewidth=0.5)
ax3.set_title('Emission Rates (Both Nodes)')
ax3.set_xlabel('Time (hours)')
ax3.set_ylabel('Emission Rate [kg CO$_2$/MWh]')
ax3.set_xlim(0, 8759)
ax3.set_ylim(0, 300)
ax3.legend(loc='upper right')

# Save or show
if save:
    fig1.savefig(f"{savepath}eprice_hourly_profiles.pdf", format='pdf')
plt.show()

# === FIGURE 2: ECDF PLOTS WITH SWAPPED AXES ===
fig2, axes2 = plt.subplots(1, 2, figsize=(7, 3.5), sharey=True)
fig2.subplots_adjust(wspace=0.3, top=0.9, bottom=0.15, left=0.1, right=0.98)

# ECDF Price (x=probability, y=price)
ax_price = axes2[0]
for node, colors in zip(nodes, [colors_chemelot, colors_zeeland]):
    for i, year in enumerate(years):
        prices = data[node][year]['price'].sort_values()
        prob = np.linspace(0, 1, len(prices))
        ax_price.plot(prob, prices, color=colors[i], label=f'{node} {labels[year]}', linewidth=0.8)
ax_price.set_title('Electricity Price')
ax_price.set_xlabel('Cumulative Probability')
ax_price.set_ylabel('Electricity Price [EUR/MWh]')
ax_price.set_xlim(0, 1)
ax_price.set_ylim(0, 200)
ax_price.legend(loc='upper left')

# ECDF Emission (x=probability, y=emission)
ax_emiss = axes2[1]
for i, year in enumerate(years):
    emissions = data['Chemelot'][year]['emission'].sort_values()
    prob = np.linspace(0, 1, len(emissions))
    ax_emiss.plot(prob, emissions, color=colors_emission[i],
                  linestyle='--', label=f'{labels[year]}', linewidth=0.8)
ax_emiss.set_title('Emission Rate')
ax_emiss.set_xlabel('Cumulative Probability')
ax_emiss.set_ylabel('Emission Rate [kg CO$_2$/MWh]')
ax_emiss.set_xlim(0, 1)
ax_emiss.set_ylim(0, 300)
ax_emiss.legend(loc='upper left')

# Save or show
# savepath = 'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/'
if save:
    fig2.savefig(f"{savepath}eprice_cumulative_dist.pdf", format='pdf')
plt.show()


# # Create the figure and layout
# fig = plt.figure(figsize=(7, 7))  # Adjust width and height
# gs = fig.add_gridspec(4, 2, height_ratios=[1.5, 1.5, 1.5, 1], width_ratios=[1, 1])
# fig.subplots_adjust(hspace=0.35, wspace=0.3, top=0.95, bottom=0.1, left=0.1, right=0.98)
#
# # Custom legend handles
# custom_lines_chemelot = [Line2D([0], [0], color=colors_chemelot[i], linewidth=2) for i in range(len(years))]
# custom_lines_zeeland = [Line2D([0], [0], color=colors_zeeland[i], linewidth=2) for i in range(len(years))]
# custom_lines_emission = [Line2D([0], [0], color=colors_emission[i], linewidth=2) for i in range(len(years))]
#
# # Row 1: Price plot for Chemelot
# ax_price_chemelot = fig.add_subplot(gs[0, :])
# avg_prices_chemelot = {year: data['Chemelot'][year]['price'].mean() for year in years}
# for i, year in enumerate(years):
#     ax_price_chemelot.plot(data['Chemelot'][year]['price'], color=colors_chemelot[i], label=f"{labels[year]} ({avg_prices_chemelot[year]:.1f} EUR/MWh)", linewidth=0.5)
# ax_price_chemelot.set_title('Electricity Prices Chemelot')
# ax_price_chemelot.set_ylabel('Electricity Price\n[EUR/MWh]')
# ax_price_chemelot.set_xlim(0, 8759)
# ax_price_chemelot.set_ylim(0, 200)
# ax_price_chemelot.legend(custom_lines_chemelot, [f"{labels[year]} ({avg_prices_chemelot[year]:.1f} EUR/MWh)" for year in years], loc='upper right')
#
# # Row 2: Price plot for Zeeland
# ax_price_zeeland = fig.add_subplot(gs[1, :])
# avg_prices_zeeland = {year: data['Zeeland'][year]['price'].mean() for year in years}
# for i, year in enumerate(years):
#     ax_price_zeeland.plot(data['Zeeland'][year]['price'], color=colors_zeeland[i], label=f"{labels[year]} ({avg_prices_zeeland[year]:.1f} EUR/MWh)", linewidth=0.5)
# ax_price_zeeland.set_title('Electricity Prices: Zeeland')
# ax_price_zeeland.set_ylabel('Electricity Price\n[EUR/MWh]')
# ax_price_zeeland.set_xlim(0, 8759)
# ax_price_zeeland.set_ylim(0, 200)
# ax_price_zeeland.legend(custom_lines_zeeland, [f"{labels[year]} ({avg_prices_zeeland[year]:.1f} EUR/MWh)" for year in years], loc='upper right')
#
# # Row 3: Emission rate (common for both nodes)
# ax_emission = fig.add_subplot(gs[2, :])
# avg_emission = {year: data['Chemelot'][year]['emission'].mean() for year in years}
# for i, year in enumerate(years):
#     ax_emission.plot(data['Chemelot'][year]['emission'], color=colors_emission[i], linestyle='-', label=f"{labels[year]} ({avg_emission[year]:.0f} [kg CO$_2$/MWh])", linewidth=0.5)
# ax_emission.set_title('Emission Rates (Both Nodes)')
# ax_emission.set_xlabel('Time (hours)')
# ax_emission.set_ylabel('Emission Rate\n[kg CO$_2$/MWh]')
# ax_emission.set_xlim(0, 8759)
# # ax_emission.set_ylim(0, 0.3)
# ax_emission.set_ylim(0, 300)
# ax_emission.legend(custom_lines_emission, [f"{labels[year]} ({avg_emission[year]:.0f} [kg CO$_2$/MWh])" for year in years], loc='upper right')
#
# # Row 4a: ECDF for electricity prices (Chemelot + Zeeland)
# ax_hist_price = fig.add_subplot(gs[3, 0])
# for node, colors in zip(nodes, [colors_chemelot, colors_zeeland]):
#     for i, year in enumerate(years):
#         ecdf_price = pd.Series(data[node][year]['price']).value_counts(normalize=True).sort_index().cumsum()
#         ax_hist_price.plot(ecdf_price.index, 1 - ecdf_price.values, color=colors[i], label=f'{node} {labels[year]}', linewidth=0.8)
# ax_hist_price.set_title('Electricity Prices')
# ax_hist_price.set_xlabel('Electricity Price [EUR/MWh]')
# ax_hist_price.set_ylabel('Probability')
# ax_hist_price.set_xlim(0, 200)
# # ax_hist_price.legend(loc='upper right')
#
# # Row 4b: ECDF for emissions (Chemelot + Zeeland)
# ax_hist_emission = fig.add_subplot(gs[3, 1])
# for i, year in enumerate(years):
#     ecdf_emission = pd.Series(data[node][year]['emission']).value_counts(normalize=True).sort_index().cumsum()
#     ax_hist_emission.plot(ecdf_emission.index, 1 - ecdf_emission.values, color=colors_emission[i], linestyle='--', label=f'{node} {labels[year]}', linewidth=0.8)
# ax_hist_emission.set_title('Emission Rates')
# ax_hist_emission.set_xlabel('Emission Rate [kg CO$_2$/MWh]')
# ax_hist_emission.set_ylabel('Probability')
# # ax_hist_emission.legend(loc='upper right')
#
# # Save and show plot
# savepath = 'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/'
# plt.savefig(f"{savepath}eprice_combined.pdf", format='pdf')
# plt.show()
