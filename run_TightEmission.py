import json
from pathlib import Path
import pandas as pd
import adopt_net0.data_preprocessing as dp
from adopt_net0.modelhub import ModelHub
from adopt_net0.result_management.read_results import add_values_to_summary
from adopt_net0.utilities import fix_installed_capacities, installed_capacities_existing, \
    installed_capacities_existing_from_file
import os

#Define basepath
basepath = os.path.dirname(os.path.abspath(__file__))

#Run Chemelot cluster case brownfield
execute = 1

if execute == 1:
    # Specify the base path to your input data
    casepath = os.path.join(basepath, "Case_studies", "MY_Chemelot_bf_")
    resultpath = os.path.join(basepath, "Raw_results", "EmissionLimit Brownfield")

    # select simulation types
    sensitivity = 'TightEmission'
    scope3 = 1
    run_with_emission_limit = 1
    intervals = ['2040', '2050']
    # intervals = ['2030', '2040', '2050']
    interval_emissionLim = {'2030': 0.75, '2040': 0.5, '2050': 0}
    nr_DD_days = 10
    prev_from_file = 1
    emission_2030 = 522537.2155
    h5_path_prev = Path(
        "Z:/AdOpt_NET0/AdOpt_results/MY/EmissionLimit Brownfield/TightEmission/20250422114955_2030_tight_minC_DD10-1/optimization_results.h5")
    pyhub = {}

    for i, interval in enumerate(intervals):
        casepath_interval = casepath + interval
        json_filepath = Path(casepath_interval) / "ConfigModel.json"

        with open(json_filepath) as json_file:
            model_config = json.load(json_file)

        model_config['optimization']['typicaldays']['N']['value'] = nr_DD_days
        model_config['optimization']['objective']['value'] = "costs_emissionlimit"

        # Set emission limit based on original 2030 case
        limit = interval_emissionLim[interval] * emission_2030
        model_config['optimization']['emission_limit']['value'] = limit

        # Scope 3 analysis yes/no
        model_config['optimization']['scope_three_analysis'] = scope3

        # solver settings
        model_config['solveroptions']['timelim']['value'] = 24*30
        model_config['solveroptions']['mipgap']['value'] = 0.01
        model_config['solveroptions']['threads']['value'] = 12
        model_config['solveroptions']['nodefilestart']['value'] = 200

        # change save options
        model_config['reporting']['save_summary_path']['value'] = resultpath + sensitivity
        model_config['reporting']['save_path']['value'] = resultpath + sensitivity

        # Write the updated JSON data back to the file
        with open(json_filepath, 'w') as json_file:
            json.dump(model_config, json_file, indent=4)

        if interval != '2030':
            prev_interval = intervals[i - 1]
            if prev_from_file and interval == '2040':
                if h5_path_prev.exists():
                    installed_capacities_existing_from_file(interval, '2030', 'Chemelot', casepath_interval, h5_path_prev)
            else:
                installed_capacities_existing(pyhub, interval, prev_interval, 'Chemelot', casepath_interval)

        # Construct and solve the model
        if interval == '2030':
            interval = interval + '_tight'
        pyhub[interval] = ModelHub()
        pyhub[interval].read_data(casepath_interval)

        # Set case name
        if nr_DD_days > 0:
            pyhub[interval].data.model_config['reporting']['case_name'][
                'value'] = (interval + '_minC_' +
                            'DD' + str(pyhub[interval].data.model_config['optimization']['typicaldays']['N']['value']))

        else:
            pyhub[interval].data.model_config['reporting']['case_name'][
                'value'] = interval + '_minC_fullres'

        # Start brownfield optimization
        pyhub[interval].construct_model()
        pyhub[interval].construct_balances()
        pyhub[interval].solve()

        # if interval == '2030':
        #     casepath_interval = casepath + interval
        #     json_filepath = Path(casepath_interval) / "ConfigModel.json"
        #
        #     with open(json_filepath) as json_file:
        #         model_config = json.load(json_file)
        #
        #     model_config['optimization']['objective']['value'] = "costs_emissionlimit"
        #
        #     if nr_DD_days > 0:
        #         limit = interval_emissionLim[interval] * pyhub[interval].model[
        #             'clustered'].var_emissions_net.value
        #     else:
        #         limit = interval_emissionLim[interval] * pyhub[interval].model['full'].var_emissions_net.value
        #     model_config['optimization']['emission_limit']['value'] = limit
        #
        #     # Scope 3 analysis yes/no
        #     model_config['optimization']['scope_three_analysis'] = scope3
        #
        #     # solver settings
        #     model_config['solveroptions']['timelim']['value'] = 24 * 30
        #     model_config['solveroptions']['mipgap']['value'] = 0.01
        #     model_config['solveroptions']['threads']['value'] = 12
        #     model_config['solveroptions']['nodefilestart']['value'] = 200
        #
        #     # change save options
        #     model_config['reporting']['save_summary_path']['value'] = resultpath + sensitivity
        #     model_config['reporting']['save_path']['value'] = resultpath + sensitivity
        #
        #     # Write the updated JSON data back to the file
        #     with open(json_filepath, 'w') as json_file:
        #         json.dump(model_config, json_file, indent=4)
        #
        #     # Construct and solve the model
        #     interval_tight = interval + '_tight'
        #     pyhub[interval_tight] = ModelHub()
        #     pyhub[interval_tight].read_data(casepath_interval)
        #
        #     # Set case name
        #     if nr_DD_days > 0:
        #         pyhub[interval_tight].data.model_config['reporting']['case_name'][
        #             'value'] = (interval_tight + '_minC_' +
        #                         'DD' + str(
        #                     pyhub[interval_tight].data.model_config['optimization']['typicaldays']['N']['value']))
        #
        #     else:
        #         pyhub[interval].data.model_config['reporting']['case_name'][
        #             'value'] = interval_tight + '_minC_fullres'
        #
        #     # Start brownfield optimization
        #     pyhub[interval_tight].construct_model()
        #     pyhub[interval_tight].construct_balances()
        #     pyhub[interval_tight].solve()


