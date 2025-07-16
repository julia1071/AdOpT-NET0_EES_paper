import os
from pathlib import Path

import h5py
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from scipy.interpolate import make_interp_spline

from adopt_net0 import extract_datasets_from_h5group

# Define the data paths
# RESULT_FOLDER = os.path.join(basepath, "Raw_results", "EmissionScope Brownfield")
RESULT_FOLDER = os.path.join(basepath, "Raw_results", "EmissionLimit Brownfield")
# DATA_TO_EXCEL_PATH1 = 'C:/Users/5637635/PycharmProjects/AdOpT-NET0_Julia//Plotting/production_shares_olefins_scope.xlsx'
# DATA_TO_EXCEL_PATH2 = 'C:/Users/5637635/PycharmProjects/AdOpT-NET0_Julia//Plotting/production_shares_ammonia_scope.xlsx'
DATA_TO_EXCEL_PATH1 = 'C:/Users/5637635/PycharmProjects/AdOpT-NET0_Julia//Plotting/production_shares_olefins.xlsx'
DATA_TO_EXCEL_PATH2 = 'C:/Users/5637635/PycharmProjects/AdOpT-NET0_Julia//Plotting/production_shares_ammonia.xlsx'
DATAPATH = "C:/Users/5637635/PycharmProjects/AdOpT-NET0_Julia/Plotting"


def fetch_and_process_data_production(resultfolder, data_to_excel_path_olefins, data_to_excel_path_ammonia,
                                      sensitivities, result_type, tec_mapping, categories):
    all_results = []
    olefin_results = []
    ammonia_results = []

    for sensitivity in sensitivities:
        resultfolder_type = f"{resultfolder}{sensitivity}"

        columns = pd.MultiIndex.from_product(
            [[result_type], [str(sensitivity)], ['2030', '2040', '2050']],
            names=["Resulttype", "Location", "Interval"]
        )

        result_data = pd.DataFrame(0.0, index=tec_mapping.keys(), columns=columns)
        production_sum_olefins = pd.DataFrame(0.0, index=categories.keys(), columns=columns)
        production_sum_ammonia = pd.DataFrame(0.0, index=categories.keys(), columns=columns)

        #read summary data
        summarypath = os.path.join(resultfolder_type, "Summary.xlsx")

        try:
            summary_results = pd.read_excel(summarypath)
        except FileNotFoundError:
            print(f"Warning: Summary file not found for {sensitivity} - {sensitivity}")
            continue

        for interval in result_data.columns.levels[2]:
            for case in summary_results['case']:
                if pd.notna(case) and interval in case:
                    h5_path = Path(summary_results[summary_results['case'] == case].iloc[0][
                                       'time_stamp']) / "optimization_results.h5"
                    if h5_path.exists():
                        with h5py.File(h5_path, "r") as hdf_file:
                            tec_operation = extract_datasets_from_h5group(
                                hdf_file["operation/technology_operation"])
                            tec_operation = {k: v for k, v in tec_operation.items() if len(v) >= 8670}
                            df_tec_operation = pd.DataFrame(tec_operation)
                            if sensitivity in ["MPWemission", "OptBIO", "noCO2electrolysis"]:
                                sensitivity_data = "Chemelot"
                            else:
                                sensitivity_data = sensitivity

                            for tec in tec_mapping.keys():
                                para = tec_mapping[tec][2] + "_output"
                                if (interval, sensitivity_data, tec, para) in df_tec_operation:
                                    output_car = df_tec_operation[interval, sensitivity_data, tec, para]

                                    if tec in ['CrackerFurnace', 'MPW2methanol', 'SteamReformer'] and (
                                            interval, sensitivity_data, tec, 'CO2captured_output') in df_tec_operation:
                                        numerator = df_tec_operation[
                                            interval, sensitivity_data, tec, 'CO2captured_output'].sum()
                                        denominator = (
                                                df_tec_operation[
                                                    interval, sensitivity_data, tec, 'CO2captured_output'].sum()
                                                + df_tec_operation[interval, sensitivity_data, tec, 'emissions_pos'].sum()
                                        )

                                        frac_CC = numerator / denominator if (
                                                denominator > 1 and numerator > 1) else 0

                                        tec_CC = tec + "_CC"
                                        result_data.loc[tec, (result_type, sensitivity, interval)] = sum(
                                            output_car) * (1 - frac_CC)
                                        result_data.loc[tec_CC, (result_type, sensitivity, interval)] = sum(
                                            output_car) * frac_CC
                                    else:
                                        result_data.loc[tec, (result_type, sensitivity, interval)] = sum(output_car)

                                tec_existing = tec + "_existing"
                                if (interval, sensitivity_data, tec_existing, para) in df_tec_operation:
                                    output_car = df_tec_operation[interval, sensitivity_data, tec_existing, para]

                                    if tec in ['CrackerFurnace', 'MPW2methanol', 'SteamReformer'] and (
                                            interval, sensitivity_data, tec_existing,
                                            'CO2captured_output') in df_tec_operation:
                                        numerator = df_tec_operation[
                                            interval, sensitivity_data, tec_existing, 'CO2captured_output'].sum()
                                        denominator = (
                                                df_tec_operation[
                                                    interval, sensitivity_data, tec_existing, 'CO2captured_output'].sum()
                                                + df_tec_operation[
                                                    interval, sensitivity_data, tec_existing, 'emissions_pos'].sum()
                                        )

                                        frac_CC = numerator / denominator if (
                                                denominator > 1 and numerator > 1) else 0

                                        tec_CC = tec + "_CC"
                                        result_data.loc[tec, (result_type, sensitivity, interval)] += sum(
                                            output_car) * (1 - frac_CC)
                                        result_data.loc[tec_CC, (result_type, sensitivity, interval)] += sum(
                                            output_car) * frac_CC
                                    else:
                                        result_data.loc[tec, (result_type, sensitivity, interval)] += sum(output_car)

                        for tec in tec_mapping.keys():
                            if tec_mapping[tec][0] == 'Olefin':
                                olefin_production = result_data.loc[tec, (result_type, sensitivity, interval)] * \
                                                    tec_mapping[tec][3]
                                production_sum_olefins.loc[
                                    tec_mapping[tec][1], (result_type, sensitivity, interval)] += olefin_production
                            elif tec_mapping[tec][0] == 'Ammonia':
                                ammonia_production = result_data.loc[tec, (result_type, sensitivity, interval)] * \
                                                     tec_mapping[tec][3]
                                production_sum_ammonia.loc[
                                    tec_mapping[tec][1], (result_type, sensitivity, interval)] += ammonia_production

        all_results.append(result_data)
        olefin_results.append(production_sum_olefins)
        ammonia_results.append(production_sum_ammonia)

    production_sum_olefins = pd.concat(olefin_results, axis=1)
    production_sum_olefins.to_excel(data_to_excel_path_olefins)
    production_sum_ammonia = pd.concat(ammonia_results, axis=1)
    production_sum_ammonia.to_excel(data_to_excel_path_ammonia)


def plot_production_shares(df, categories):
    plt.rcParams.update({'font.family': 'serif', 'font.size': 14})

    df.columns.name = None
    df = df.T.reset_index()
    df = df.rename(columns={'index': 'Year'})
    df['Year'] = df['Year'].astype(int)

    years = df['Year'].values
    available_categories = [cat for cat in categories if cat in df.columns]
    df = df[available_categories]

    shares = df.div(df.sum(axis=1), axis=0)
    x_smooth = np.linspace(years.min(), years.max(), 300)

    interpolated = {}
    for col in available_categories:
        spline = make_interp_spline(years, shares[col], k=2)
        interpolated[col] = spline(x_smooth)

    y_stack = np.row_stack([interpolated[col] for col in available_categories])

    fig, ax = plt.subplots(figsize=(6, 3))
    ax.stackplot(x_smooth, y_stack, labels=available_categories, colors=[categories[c] for c in available_categories])
    # ax.set_title("Share of Total Production by Technology")
    ax.set_ylabel("Share of Total Production")
    ax.set_xlabel("Year")
    ax.legend(loc='upper left', bbox_to_anchor=(1, 1))
    ax.set_xlim(years.min(), years.max())
    ax.set_xticks([2030, 2040, 2050])
    ax.set_xticklabels([r"$2030$", r"$2040$", r"$2050$"])
    ax.set_ylim(0, 1)
    plt.rcParams['font.family'] = 'serif'
    plt.tight_layout()


def plot_production_shares_stacked(df1, df2, categories, interpolation="spline", separate=1):
    plt.rcParams.update({'font.family': 'serif', 'font.size': 14})

    def preprocess(df):
        df.columns.name = None
        df = df.T.reset_index()
        df = df.rename(columns={'index': 'Year'})
        df['Year'] = df['Year'].astype(int)
        return df

    df1 = preprocess(df1)
    df2 = preprocess(df2)

    all_years = sorted(set(df1['Year']) | set(df2['Year']) | {2025})

    def fill_missing_years(df, years):
        for y in years:
            if y not in df['Year'].values:
                row = {cat: 0 for cat in categories}
                row['Conventional'] = 1
                row['Year'] = y
                df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        return df.sort_values('Year')

    df1 = fill_missing_years(df1, all_years)
    df2 = fill_missing_years(df2, all_years)

    def compute_shares(df):
        df_cat = df[[c for c in categories if c in df.columns]]
        return df_cat.div(df_cat.sum(axis=1), axis=0)

    def interpolate(df, years):
        shares = compute_shares(df)
        if interpolation == "spline":
            x = np.linspace(min(years), max(years), 300)
            interpolated = {
                col: make_interp_spline(years, shares[col], k=2)(x)
                for col in shares.columns
            }
        elif interpolation == "linear":
            x = np.array(years)
            interpolated = {
                col: np.interp(x, years, shares[col])
                for col in shares.columns
            }
        elif interpolation == "step":
            interpolated = {}

            x = []
            for i, year in enumerate(years):
                if i == 0:
                    x.append(year)
                else:
                    x.extend([year - 1, year])
            x.append(2055)  # Add final point
            x = np.array(x)

            for col in shares.columns:
                y = np.append(shares[col].values, shares[col].values[-1])
                y_step = np.repeat(y, 2)[:-1]

                # Trim or pad y_step to match x
                if len(y_step) > len(x):
                    y_step = y_step[:len(x)]
                elif len(y_step) < len(x):
                    y_step = np.append(y_step, [y_step[-1]] * (len(x) - len(y_step)))

                y_interp = y_step.copy()

                for i in range(1, len(x), 2):
                    if x[i] - x[i - 1] == 1:
                        y0, y1 = y_step[i - 1], y_step[i]
                        t = np.linspace(0, 1, 2)
                        spline = make_interp_spline(t, [y0, y1], k=2)
                        y_interp[i - 1] = spline(0)
                        y_interp[i] = spline(1)

                interpolated[col] = y_interp

        else:
            raise ValueError(f"Unsupported interpolation method: {interpolation}")

        return x, interpolated

    if separate == 1:
        fig, (ax1, ax2) = plt.subplots(nrows=2, figsize=(7, 5.5), sharex=True,
                                       gridspec_kw={'hspace': 0.1}
                                       )

        for ax, df, label in zip((ax1, ax2), (df1, df2), ('ammonia', 'ethylene')):
            x, interpolated = interpolate(df, df['Year'].values)
            bottoms = np.zeros_like(x)
            for cat in categories:
                if cat in interpolated:
                    top = bottoms + interpolated[cat]
                    ax.fill_between(x, bottoms, top, color=categories[cat], label=cat)
                    bottoms = top
            ax.set_ylim(0, 1)
            ax.set_ylabel(f"Share of {label}")

        ax2.set_xticks([2025, 2030, 2040, 2050, 2055])
        ax2.set_xticklabels([r"Current", 2030, 2040, 2050, r"Post 2050"])
        ax2.set_xlim(x.min(), x.max())

        # Combine legend from both axes
        # handles, labels = ax2.get_legend_handles_labels()
        # fig.legend(handles, labels,
        #            loc='lower center',
        #            bbox_to_anchor=(0.5, 0),
        #            ncol=2)
        # plt.subplots_adjust(bottom=0.25)

    else:
        # Merge and plot together
        merged = df1.copy()
        for cat in categories:
            merged[cat] = df1.get(cat, 0) + df2.get(cat, 0)
        merged = fill_missing_years(merged, all_years)
        x, interpolated = interpolate(merged, merged['Year'].values)

        fig, ax1 = plt.subplots(figsize=(7, 3))
        bottoms = np.zeros_like(x)
        for cat in categories:
            if cat in interpolated:
                top = bottoms + interpolated[cat]
                ax1.fill_between(x, bottoms, top, color=categories[cat], label=cat)
                bottoms = top
        ax1.set_ylabel("Share of Total Production")
        ax1.set_ylim(0, 1)
        ax1.set_xticks([2025, 2030, 2040, 2050, 2055])
        ax1.set_xticklabels([r"Current", 2030, 2040, 2050, r"Post 2050"])
        ax1.legend(loc='upper left', bbox_to_anchor=(1, 1))
        ax1.set_title("Combined Production Shares")

    plt.tight_layout()


def save_separate_legend(categories, filename="legend.pdf"):
    import matplotlib.patches as mpatches

    plt.rcParams.update({'font.family': 'serif', 'font.size': 12})
    fig, ax = plt.subplots(figsize=(6.5, 0.6))  # Adjust width/height as needed

    # Create dummy handles
    handles = [mpatches.Patch(color=color, label=label)
               for label, color in categories.items()]

    # Add the legend to the figure
    legend = fig.legend(handles, categories.keys(),
                        loc='center',
                        ncol=4,  # adjust based on layout needs
                        frameon=False)

    ax.axis('off')  # Hide axes completely

    plt.tight_layout()
    fig.savefig(f'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/{filename}', format='pdf', bbox_inches='tight')
    plt.close(fig)


def main():
    result_type = 'EmissionLimit Brownfield'
    set_sensitivities = ['Chemelot', 'Zeeland', 'noCO2electrolysis', 'MPWemission', 'OptBIO']
    # result_type = 'EmissionScope Brownfield'
    # set_sensitivities = ['Chemelot']

    tec_mapping = {
        "CrackerFurnace": ("Olefin", "Conventional", "olefins", 0.439),
        "CrackerFurnace_CC": ("Olefin", "Carbon Capture", "olefins", 0.439),
        "CrackerFurnace_Electric": ("Olefin", "Electrification", "olefins", 0.439),
        "SteamReformer": ("Ammonia", "Conventional", "HBfeed", 0.168),
        "SteamReformer_CC": ("Ammonia", "Carbon Capture", "HBfeed", 0.168),
        "WGS_m": ("Ammonia", "Electrification", "hydrogen", 0.168),
        "AEC": ("Ammonia", "Water electrolysis", "hydrogen", 0.168),
        "RWGS": ("Olefin", r"CO$_2$ utilization", "syngas", 0.270),
        "DirectMeOHsynthesis": ("Olefins", r"CO$_2$ utilization", "methanol", 0.328),
        "EDH": ("Olefin", "Bio-based feedstock", "ethylene", 1),
        "PDH": ("Olefin", "Bio-based feedstock", "propylene", 1),
        "MPW2methanol": ("Olefin", "Plastic waste recycling", "methanol", 0.328),
        "MPW2methanol_CC": ("Olefin", "Plastic waste recycling with CC", "methanol", 0.328),
        "CO2electrolysis": ("Olefin", r"CO$_2$ utilization", "ethylene", 1),
    }

    categories = {
        "Conventional": '#8C8B8B',
        "Carbon Capture": '#3E7EB0',
        "Electrification": '#E9E46D',
        "Water electrolysis": '#EABF37',
        r"CO$_2$ utilization": '#E18826',
        "Bio-based feedstock": '#84AA6F',
        "Plastic waste recycling": '#B475B2',
        "Plastic waste recycling with CC": '#533A8C',
    }

    get_data = 0

    if get_data == 1:
        fetch_and_process_data_production(RESULT_FOLDER, DATA_TO_EXCEL_PATH1, DATA_TO_EXCEL_PATH2, set_sensitivities, result_type,
                                          tec_mapping, categories)

    production_sum_olefins = pd.read_excel(DATA_TO_EXCEL_PATH1, index_col=0, header=[0, 1, 2])
    production_sum_ammonia = pd.read_excel(DATA_TO_EXCEL_PATH2, index_col=0, header=[0, 1, 2])

    product = "stacked"
    interpolation = "step"
    separate = 1

    for sensitivity in set_sensitivities:
        if product == "Olefin":
            df_plot = production_sum_olefins.loc[:, (result_type, sensitivity)].copy()
            plot_production_shares(df_plot, categories)
        elif product == 'Ammonia':
            df_plot = production_sum_ammonia.loc[:, (result_type, sensitivity)].copy()
            plot_production_shares(df_plot, categories)
        elif product == 'stacked':
            df_plot_olefin = production_sum_olefins.loc[:, (result_type, sensitivity)].copy()
            df_plot_ammonia = production_sum_ammonia.loc[:, (result_type, sensitivity)].copy()
            plot_production_shares_stacked(df_plot_ammonia, df_plot_olefin, categories, interpolation=interpolation,
                                           separate=separate)

        #Make the plots
        if 'EmissionLimit' in result_type:
            ext_map = {'Brownfield': '_bf', 'Greenfield': '_gf'}
            ext = next((v for k, v in ext_map.items() if k in result_type), '')
        else:
            ext = '_bf_scope'

        filename = f'production_share_{sensitivity}{ext}'


        saveas = 'pdf'
        if saveas == 'svg':
            savepath = f'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/{filename}.svg'
            plt.savefig(savepath, format='svg')
        elif saveas == 'pdf':
            savepath = f'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/{filename}.pdf'
            plt.savefig(savepath, format='pdf', bbox_inches='tight')
        elif saveas == 'both':
            savepath = f'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/{filename}.pdf'
            plt.savefig(savepath, format='pdf')
            savepath = f'C:/Users/5637635/OneDrive - Universiteit Utrecht/Research/Multiyear Modeling/MY_Plots/{filename}.svg'
            plt.savefig(savepath, format='svg', bbox_inches='tight')

        plt.show()

        # After all plots:
        # save_separate_legend(categories)


if __name__ == "__main__":
    main()
