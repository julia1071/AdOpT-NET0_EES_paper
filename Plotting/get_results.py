import os
import h5py
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from matplotlib.ticker import PercentFormatter
from adopt_net0 import extract_datasets_from_h5group
import os

#Add basepath
basepath = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

#options
sensitivity = 1
zeeland = 0

if sensitivity:
    data_to_excel_path = os.path.join(basepath, "Plotting", "result_data_long_sensitivity.xlsx")
    result_types = ['EmissionLimit Greenfield', 'EmissionLimit Brownfield'] # Add multiple result types
elif zeeland:
    data_to_excel_path = os.path.join(basepath, "Plotting", "result_data_long_Zeeland.xlsx")
    result_types = ['EmissionLimit Greenfield', 'EmissionLimit Brownfield']  # Add multiple result types
else:
    data_to_excel_path = os.path.join(basepath, "Plotting", "result_data_long.xlsx")
    result_types = ['EmissionLimit Greenfield', 'EmissionLimit Brownfield', 'EmissionScope Greenfield',
                    'EmissionScope Brownfield']


# Initialize an empty dictionary to collect DataFrame results
all_results = []

for result_type in result_types:
    resultfolder = os.path.join(basepath, "Plotting", f"{result_type}")

    # Define the multi-level index for rows
    if sensitivity:
        columns = pd.MultiIndex.from_product(
            [
                [str(result_type)],
                ["MPWemission", "OptBIO", "noCO2electrolysis", "TightEmission"],
                ["2030", "2040", "2050"]
            ],
            names=["Resulttype", "Location", "Interval"]
        )
    elif zeeland:
        columns = pd.MultiIndex.from_product(
            [
                [str(result_type)],
                ["Zeeland"],
                ["2030", "2040", "2050"]
            ],
            names=["Resulttype", "Location", "Interval"]
        )
    else:
        columns = pd.MultiIndex.from_product(
            [
                [str(result_type)],
                ["Chemelot"],
                ["2030", "2040", "2050"]
            ],
            names=["Resulttype", "Location", "Interval"]
        )

    # Define the index rows
    index = ["path", "costs_obj_interval", "sunk_costs", "costs_tot_interval", "costs_tot_cumulative", "emissions_net"]

    # Create the DataFrame for this result type with NaN values
    result_data = pd.DataFrame(np.nan, index=index, columns=columns)

    # Fill the path column using a loop
    for location in result_data.columns.levels[1]:
        folder_name = f"{location}"
        summarypath = os.path.join(resultfolder, folder_name, "Summary.xlsx")

        try:
            summary_results = pd.read_excel(summarypath)
        except FileNotFoundError:
            print(f"Warning: Summary file not found for {result_type} - {location}")
            continue

        tec_costs = {}
        total_costs = {}
        for case in summary_results['case']:
            for i, interval in enumerate(result_data.columns.levels[2]):
                # if sensitivity  and interval == '2030':
                #     interval = interval + '_tight'
                if pd.notna(case) and interval in case:
                    h5_path = Path(summary_results.loc[summary_results['case'] == case, 'time_stamp'].iloc[
                                       0]) / "optimization_results.h5"
                    result_data.at["path", (result_type, location, interval)] = h5_path
                    result_data.loc["costs_obj_interval", (result_type, location, interval)] = \
                        summary_results.loc[summary_results['case'] == case, 'total_npv'].iloc[0]
                    result_data.loc["emissions_net", (result_type, location, interval)] = \
                        summary_results.loc[summary_results['case'] == case, 'emissions_net'].iloc[0]
                    tec_costs[interval] = summary_results.loc[summary_results['case'] == case, 'cost_capex_tecs'].iloc[0]
                    total_costs[interval] = summary_results.loc[summary_results['case'] == case, 'total_npv'].iloc[
                        0]

                    #Calculate sunk costs and cumulative costs for brownfield
                    if 'Brownfield' in result_type:
                        prev_interval = result_data.columns.levels[2][i - 1]
                        if interval == '2030':
                            result_data.loc["costs_tot_interval", (result_type, location, interval)] = \
                                summary_results.loc[summary_results['case'] == case, 'total_npv'].iloc[0]
                        if interval == '2040':
                            result_data.loc["sunk_costs", (result_type, location, interval)] = tec_costs[prev_interval]
                            result_data.loc["costs_tot_interval", (result_type, location, interval)] = tec_costs[prev_interval] + \
                                summary_results.loc[summary_results['case'] == case, 'total_npv'].iloc[0]
                        if interval == '2050':
                            first_interval = result_data.columns.levels[2][i - 2]
                            result_data.loc["sunk_costs", (result_type, location, interval)] = tec_costs[
                                prev_interval] + tec_costs[first_interval]
                            result_data.loc["costs_tot_interval", (result_type, location, interval)] = tec_costs[prev_interval] + \
                                + tec_costs[first_interval] + summary_results.loc[summary_results['case'] == case, 'total_npv'].iloc[0]
                            result_data.loc["costs_tot_cumulative", (result_type, location, interval)] = sum(
                                total_costs.values()) * 10 + tec_costs[prev_interval] * 10 + tec_costs[first_interval] * 10


                    # Calculate total cumulative costs for Greenfield
                    if 'Greenfield' in result_type:
                        result_data.loc["costs_tot_cumulative", (result_type, location, interval)] = total_costs[interval] * 30

                    if h5_path.exists():
                        with h5py.File(h5_path, "r") as hdf_file:
                            nodedata = extract_datasets_from_h5group(hdf_file["design/nodes"])
                            df_nodedata = pd.DataFrame(nodedata)

                            if sensitivity:
                                location_data = 'Chemelot'
                                interval_data = interval
                            else:
                                location_data = location
                                interval_data = interval

                            for tec in df_nodedata.columns.levels[2]:
                                output_name = f'size_{tec}'
                                if (interval_data, location_data, tec, 'size') in df_nodedata.columns:
                                    result_data.loc[output_name, (result_type, location, interval)] = \
                                        df_nodedata[(interval_data, location_data, tec, 'size')].iloc[0]
                                else:
                                    result_data.loc[output_name, (result_type, location, interval)] = 0

                                if any(tec.startswith(base) for base in ['CrackerFurnace', 'MPW2methanol', 'SteamReformer']):
                                    tec_operation = extract_datasets_from_h5group(
                                        hdf_file["operation/technology_operation"])
                                    tec_operation = {k: v for k, v in tec_operation.items() if len(v) >= 8670}
                                    df_tec_operation = pd.DataFrame(tec_operation)
                                    if (interval_data, location_data, tec, 'CO2captured_output') in df_tec_operation:
                                        numerator = df_tec_operation[
                                            interval_data, location_data, tec, 'CO2captured_output'].sum()
                                        denominator = (
                                                df_tec_operation[
                                                    interval_data, location_data, tec, 'CO2captured_output'].sum()
                                                + df_tec_operation[interval_data, location_data, tec, 'emissions_pos'].sum()
                                        )

                                        frac_CC = numerator / denominator if (denominator > 1 and numerator > 1) else 0

                                        tec_CC = "size_" + tec + "_CC"
                                        if tec_CC not in result_data.index:
                                            result_data.loc[tec_CC] = pd.Series(dtype=float)
                                        result_data.loc[tec_CC, (result_type, location, interval)] = frac_CC


                            ebalance = extract_datasets_from_h5group(hdf_file["operation/energy_balance"])
                            df_ebalance = pd.DataFrame(ebalance)
                            cars_at_node = df_ebalance[interval_data, location_data].columns.droplevel([1]).unique()

                            for car in cars_at_node:
                                parameters = ["import", "export"]
                                for para in parameters:
                                    output_name = f"{car}/{para}_max"
                                    if (interval_data, location_data, car, para) in df_ebalance.columns:
                                        car_output = df_ebalance[interval_data, location_data, car, para]
                                        result_data.loc[output_name, (result_type, location, interval)] = max(
                                            car_output)
                                    else:
                                        result_data.loc[output_name, (result_type, location, interval)] = 0


    # Store results for this result_type
    all_results.append(result_data)

# Concatenate all result types into a single DataFrame
final_result_data = pd.concat(all_results, axis=1)

# Save to Excel (optional)
final_result_data.to_excel(data_to_excel_path)
