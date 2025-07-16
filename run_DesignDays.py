import json
from pathlib import Path
import pandas as pd
import adopt_net0.data_preprocessing as dp
from adopt_net0.modelhub import ModelHub
from adopt_net0.result_management.read_results import add_values_to_summary


#Run Chemelot test design days greenfield
execute = 0

if execute == 1:
    # Specify the path to your input data
    casepath = "Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_gf_2030"
    resultpath = "Z:/AdOpt_NET0/AdOpt_results/MY/DesignDays/CH_2030_gf"

    json_filepath = Path(casepath) / "ConfigModel.json"

    node = 'Chemelot'
    scope3 = 1
    interval = '2030'
    interval_emissionLim = {'2030': 1, '2040': 0.5, '2050': 0}
    nr_DD_days = [5, 10, 20, 40, 100, 0]

    for nr in nr_DD_days:
        with open(json_filepath) as json_file:
            model_config = json.load(json_file)

        # change save options
        model_config['reporting']['save_summary_path']['value'] = resultpath
        model_config['reporting']['save_path']['value'] = resultpath

        model_config['optimization']['typicaldays']['N']['value'] = nr
        model_config['optimization']['objective']['value'] = 'costs'

        # Scope 3 analysis yes/no
        model_config['optimization']['scope_three_analysis'] = scope3

        # solver settings
        model_config['solveroptions']['timelim']['value'] = 24*30
        model_config['solveroptions']['mipgap']['value'] = 0.01
        model_config['solveroptions']['threads']['value'] = 8
        model_config['solveroptions']['nodefilestart']['value'] = 200

        # Write the updated JSON data back to the file
        with open(json_filepath, 'w') as json_file:
            json.dump(model_config, json_file, indent=4)

        # Construct and solve the model
        pyhub = ModelHub()
        pyhub.read_data(casepath)

        #add casename based on resolution
        if pyhub.data.model_config['optimization']['typicaldays']['N']['value'] == 0:
            pyhub.data.model_config['reporting']['case_name']['value'] = 'fullres'
        else:
            pyhub.data.model_config['reporting']['case_name']['value'] = 'DD' + str(
                pyhub.data.model_config['optimization']['typicaldays']['N']['value'])

        #solving
        pyhub.quick_solve()


#Run Chemelot test design days brownfield
execute = 1

if execute == 1:
    # Specify the path to your input data
    casepath = "Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_bf_2030"
    resultpath = "Z:/AdOpt_NET0/AdOpt_results/MY/DesignDays/CH_2030_bf"

    json_filepath = Path(casepath) / "ConfigModel.json"

    node = 'Chemelot'
    scope3 = 1
    interval = '2030'
    interval_emissionLim = {'2030': 1, '2040': 0.5, '2050': 0}
    nr_DD_days = [5, 10, 20, 40, 100, 0]

    for nr in nr_DD_days:
        with open(json_filepath) as json_file:
            model_config = json.load(json_file)

        # change save options
        model_config['reporting']['save_summary_path']['value'] = resultpath
        model_config['reporting']['save_path']['value'] = resultpath

        model_config['optimization']['typicaldays']['N']['value'] = nr
        model_config['optimization']['objective']['value'] = 'costs'

        # Scope 3 analysis yes/no
        model_config['optimization']['scope_three_analysis'] = scope3

        # solver settings
        model_config['solveroptions']['timelim']['value'] = 24*30
        model_config['solveroptions']['mipgap']['value'] = 0.01
        model_config['solveroptions']['threads']['value'] = 8
        model_config['solveroptions']['nodefilestart']['value'] = 200

        # Write the updated JSON data back to the file
        with open(json_filepath, 'w') as json_file:
            json.dump(model_config, json_file, indent=4)

        # Construct and solve the model
        pyhub = ModelHub()
        pyhub.read_data(casepath)

        #add casename based on resolution
        if pyhub.data.model_config['optimization']['typicaldays']['N']['value'] == 0:
            pyhub.data.model_config['reporting']['case_name']['value'] = 'fullres'
        else:
            pyhub.data.model_config['reporting']['case_name']['value'] = 'DD' + str(
                pyhub.data.model_config['optimization']['typicaldays']['N']['value'])

        #solving
        pyhub.quick_solve()
