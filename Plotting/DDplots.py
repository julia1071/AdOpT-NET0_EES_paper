import h5py
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from matplotlib.ticker import PercentFormatter
from adopt_net0 import extract_datasets_from_h5group
from matplotlib import gridspec

# Define the data path
run_for = 'gf'
interval = '2030'
resultfolder = "Z:/AdOpt_NET0/AdOpt_results/MY/DesignDays/CH_2030_" + run_for

# Define reference case ('fullres' or 'DD..' and technologies
# tecs = ['AEC', 'MTO', 'Storage_Ammonia']
tecs = ['AEC']
reference_case = 'DD10'

# Extract the relevant data
summarypath = Path(resultfolder) / 'Summary.xlsx'
summary_results = pd.read_excel(summarypath)

# full resolution reference
reference_case_results = {
    'costs': summary_results[summary_results['case'] == reference_case].iloc[0]['total_npv'],
    'emissions': summary_results[summary_results['case'] == reference_case].iloc[0]['emissions_pos'],
    'computation time': summary_results[summary_results['case'] == reference_case].iloc[0]['time_total']
}
reference_case_path_h5 = Path(summary_results[summary_results['case'] == reference_case].iloc[0]['time_stamp']) / "optimization_results.h5"

#get capacities from h5 files
if reference_case_path_h5.exists():
    with h5py.File(reference_case_path_h5, "r") as hdf_file:
        df = extract_datasets_from_h5group(hdf_file["design/nodes"])
        for tec in tecs:
            output_name = 'size_' + tec
            reference_case_results[output_name] = df[(interval, 'Chemelot', tec, 'size')][0]


#now collect data to plot
data = []
for case in summary_results['case']:
    if pd.notna(case) and case != reference_case:
        case_data = {
            'case': case,
            'DD': int(case[2:]),
            'costs': summary_results.loc[summary_results['case'] == case, 'total_npv'].iloc[0],
            'emissions': summary_results.loc[summary_results['case'] == case, 'emissions_pos'].iloc[0],
            'computation time': summary_results.loc[summary_results['case'] == case, 'time_total'].iloc[0]
        }

        #collect sizes from h5
        case_path_h5 = Path(
            summary_results[summary_results['case'] == case].iloc[0]['time_stamp']) / "optimization_results.h5"

        if reference_case_path_h5.exists():
            with h5py.File(reference_case_path_h5, "r") as hdf_file:
                df = extract_datasets_from_h5group(hdf_file["design/nodes"])
                for tec in tecs:
                    output_name = 'size_' + tec
                    case_data[output_name] = df[(interval, 'Chemelot', tec, 'size')][0]

        data.append(case_data)

# Convert the list of dictionaries into a DataFrame
clustered_results = pd.DataFrame(data)
clustered_results.set_index('case', inplace=True)

# difference
columns_of_interest = ['costs', 'emissions', 'computation time']
columns_of_interest.extend(['size_' + tec for tec in tecs])
for column in columns_of_interest:
    clustered_results[f'{column}_diff'] = ((clustered_results[column] - reference_case_results[column]) / reference_case_results[column]) * 100

# Configure Matplotlib to use LaTeX for text rendering and set font
plt.rcParams.update({
    "text.usetex": True,
    "font.family": "serif",
    "font.serif": ["Computer Modern Roman"],
    "axes.labelsize": 12,
    "font.size": 12,
    "legend.fontsize": 10,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
})

# Define custom colors
# colors = ['#9fa0c3', '#8b687f', '#7b435b', '#4C4D71', '#B096A7']
colors = ['#639051', 'black', '#5277B7', '#FFBC47']
markers = ['o', 'x', 'D', '^']

# Choose the plot type: 'diff' for percentage differences, 'absolute' for absolute values
plot_type = 'diff'  # Can be changed to 'diff' and 'absolute'

# Create figure with broken y-axis layout
fig = plt.figure(figsize=(5, 4))
gs = gridspec.GridSpec(2, 1, height_ratios=[1, 2], hspace=0.05)  # Top (large y), bottom (normal y)
ax2 = fig.add_subplot(gs[0])  # Top subplot for high values
ax1 = fig.add_subplot(gs[1], sharex=ax2)  # Bottom subplot for low values

# Plot scatter points on both axes
for i, column in enumerate(columns_of_interest):
    if plot_type == 'diff' or (plot_type == 'absolute' and column != 'computation time'):
        for tec in tecs:
            label = f'Size Difference {tec}' if column == f'size_{tec}' else f'{column.capitalize()} Difference'
            y_data = clustered_results[f'{column}_diff']
            x_data = clustered_results['DD']
            ax1.scatter(x_data, y_data, color=colors[i], label=label, marker=markers[i])
            ax2.scatter(x_data, y_data, color=colors[i], marker=markers[i])

# Set y-limits for broken axis
if run_for == 'bf':
    ax1.set_ylim(-100, 200)      # Bottom axis
    ax2.set_ylim(700, 1200)
else:
    ax1.set_ylim(-100, 500)  # Bottom axis
    ax2.set_ylim(2550, 2800)


# Hide tick labels on top x-axis
plt.setp(ax2.get_xticklabels(), visible=False)

# Format y-axis as percentage
ax1.yaxis.set_major_formatter(PercentFormatter(decimals=0))
ax2.yaxis.set_major_formatter(PercentFormatter(decimals=0))

# Add diagonal break marks
kwargs = dict(marker=[(-1, -1), (1, 1)], markersize=12, linestyle='none', color='k', mec='k', mew=1, clip_on=False)
ax1.plot([0, 1], [1, 1], transform=ax1.transAxes, **kwargs)
ax2.plot([0, 1], [0, 0], transform=ax2.transAxes, **kwargs)

# Axis labels
ax1.set_xlabel('Number of Design Days')
# ax1.set_ylabel('Difference with 10 Design Days (\%)')
ax2.set_ylabel('')

# Handle computation time on a secondary y-axis (only on bottom axis)
if plot_type == 'absolute':
    ax1b = ax1.twinx()
    computation_time_in_hours = clustered_results['computation time'] / 3600
    ax1b.scatter(clustered_results['DD'], computation_time_in_hours, color='black', label='Computation Time', marker='x')
    ax1b.set_ylabel('Computation Time (hours)')
    ax1b.set_ylim(-10, 100)

    # Combine legends
    handles1, labels1 = ax1.get_legend_handles_labels()
    handles2, labels2 = ax1b.get_legend_handles_labels()
    ax1.legend(handles1 + handles2, labels1 + labels2, loc='upper left')
else:
    handles, labels = ax1.get_legend_handles_labels()
    ax1.legend(handles, labels, loc='upper right')

# Add grid to both subplots
ax1.grid(True, alpha=0.2)
ax2.grid(True, alpha=0.2)

# Optional: set consistent x-ticks
xmin, xmax = ax1.get_xlim()
ax1.set_xticks(range(int(xmin), int(xmax) + 1, 10))
fig.text(0.015, 0.5, 'Difference with 10 Design Days (\%)', va='center', rotation='vertical')


# # Save and show plot
savepath = 'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/'

plt.tight_layout()
fig.subplots_adjust(left=0.15)

plt.savefig(f"{savepath}complexity_{run_for}_{plot_type}.pdf", format='pdf', bbox_inches='tight', pad_inches=0.05)
# plt.savefig(f"{savepath}complexity_{run_for}_{plot_type}.svg", format='svg', bbox_inches='tight', pad_inches=0.05)

plt.show()


# Show the plot
# plt.show()
