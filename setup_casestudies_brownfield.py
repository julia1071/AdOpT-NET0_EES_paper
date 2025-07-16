import json
import shutil
from pathlib import Path
import adopt_net0.data_preprocessing as dp
from adopt_net0.data_preprocessing.data_loading import find_json_path
from adopt_net0.modelhub import ModelHub
from adopt_net0.result_management.read_results import add_values_to_summary
import pandas as pd

#Set basepath
basepath = Path(__file__).resolve().parent

#global functions
read_all = 0

#Create data Chemelot cluster short term
execute = 1

if execute == 1:
    # Specify the path to your input data
    casepath = basepath / "Case_studies" / "MY_Chemelot_bf_2030"
    datapath = basepath / "Input_data" / "250303_MY_Data_Chemelot_bf"

    firsttime = 0
    if firsttime:
        # Create template files
        dp.create_optimization_templates(casepath)
        dp.create_montecarlo_template_csv(casepath)

        # Change topology
        topology_file = casepath / "Topology.json"
        config_file = casepath / "ConfigModel.json"

        topology_template = {
            "nodes": ["Chemelot"],
            "carriers": ["electricity", "methane", "methane_bio", "naphtha", "naphtha_bio",
                         "CO2", "CO2_DAC", "CO2captured", "hydrogen", "nitrogen", "oxygen",
                         "ammonia", "ethylene", "propylene", "PE_olefin", "olefins",
                         "crackergas", "feedgas", "steam", "heatlowT", "HBfeed", "syngas", "syngas_r",
                         "methanol", "ethanol", "propane", "MPW"],
            "investment_periods": ["2030"],
            "start_date": "2022-01-01 00:00",
            "end_date": "2022-12-31 23:00",
            "resolution": "1h",
            "investment_period_length": 1,
        }

        # open and save topology and configuration files
        with open(topology_file, "w") as f:
            json.dump(topology_template, f, indent=4)
        # with open(config_file, "w") as f:
        # json.dump(configuration_template, f, indent=4)

        # Create folder structure
        dp.create_input_data_folder_template(casepath)

        # Node location
        file_path = casepath / 'NodeLocations.csv'
        data = pd.read_csv(file_path, delimiter=';')

        data['lon'] = 5.80
        data['lat'] = 50.97
        data['alt'] = 10

        # Save the modified CSV file
        data.to_csv(file_path, index=False, sep=';')

        # Read climate data
        dp.load_climate_data_from_api(casepath)

    read_techs = 1
    if read_techs or read_all:
        set_tecs_new = ["ElectricSMR_m", "WGS_m", "AEC", "HaberBosch",
                        "CrackerFurnace_Electric",
                        "ASU", "Boiler_Industrial_NG", "Boiler_El",
                        "RWGS", "MeOHsynthesis", "MTO", "EDH", "PDH", "MPW2methanol",
                        "Storage_Ammonia", "Storage_CO2", "Storage_Ethylene",
                        "Storage_H2", "Storage_Battery", "Storage_Propylene",
                        "CO2toEmission", "feedgas_mixer", "naphtha_mixer", "PE_mixer", "CO2_mixer", "HBfeed_mixer",
                        "syngas_mixer"]
        set_tecs_existing = {"SteamReformer": 1078, "HaberBosch": 813, "CrackerFurnace": 499, "OlefinSeparation": 499}

        json_file_path = casepath / "Topology.json"
        with open(json_file_path, "r") as json_file:
            topology = json.load(json_file)

        for period in topology["investment_periods"]:
            for node_name in topology["nodes"]:
                # Read the JSON technology file
                json_tec_file_path = (
                        casepath / period / "node_data" / node_name / "Technologies.json"
                )
                with open(json_tec_file_path, "r") as json_tec_file:
                    json_tec = json.load(json_tec_file)

                json_tec['new'] = set_tecs_new
                json_tec['existing'] = set_tecs_existing
                with open(json_tec_file_path, "w") as json_tec_file:
                    json.dump(json_tec, json_tec_file, indent=4)

                # Copy technology and network data into folder
                dp.copy_technology_data(casepath, datapath)
                # Add CCS
                output_folder = (
                        casepath / period / "node_data" / node_name / "technology_data"
                )
                tec_json_file_path = find_json_path(datapath, "MEA_large")
                if tec_json_file_path:
                    shutil.copy(tec_json_file_path, output_folder)

                spec_tec_size = {"AEC": 750, "RWGS": 123, "MeOHsynthesis": 494, "MTO": 904, "SteamReformer": 1269,
                                 "ElectricSMR_m": 2114, "WGS_m": 389, "HaberBosch": 1609, "CrackerFurnace": 594,
                                 "CrackerFurnace_Electric": 386, "OlefinSeparation": 594, "ASU": 112,
                                 "Boiler_Industrial_NG": 2053, "Boiler_El": 750, "EDH": 526, "PDH": 157,
                                 "MPW2methanol": 338}
                for tec in spec_tec_size.keys():
                    tech_data_file = find_json_path(output_folder, tec)
                    with open(tech_data_file, "r") as json_tecdata_file:
                        tech_data = json.load(json_tecdata_file)

                    tech_data["size_max"] = spec_tec_size[tec]
                    with open(tech_data_file, "w") as json_tecdata_file:
                        json.dump(tech_data, json_tecdata_file, indent=4)

    read_car = 0
    if read_car or read_all:
        # Fill carrier data
        dp.fill_carrier_data(casepath, value_or_data=0)

        # Demand data
        dp.fill_carrier_data(casepath, value_or_data=135, columns=['Demand'], carriers=['ammonia'])
        dp.fill_carrier_data(casepath, value_or_data=44, columns=['Demand'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=218, columns=['Demand'], carriers=['PE_olefin'])
        dp.fill_carrier_data(casepath, value_or_data=0, columns=['Demand'], carriers=['steam'])
        dp.fill_carrier_data(casepath, value_or_data=82, columns=['Demand'], carriers=['electricity'])
        dp.fill_carrier_data(casepath, value_or_data=7, columns=['Demand'], carriers=['feedgas'])

        # Import limits
        dp.fill_carrier_data(casepath, value_or_data=750, columns=['Import limit'], carriers=['electricity'])
        dp.fill_carrier_data(casepath, value_or_data=5000, columns=['Import limit'],
                             carriers=['methane', 'naphtha', 'methane_bio', 'naphtha_bio', 'ethanol',
                                       'propane', 'MPW', 'CO2_DAC'])

        # dp.fill_carrier_data(casepath, value_or_data=2114, columns=['Import limit'], carriers=['methane'])
        # dp.fill_carrier_data(casepath, value_or_data=990, columns=['Import limit'], carriers=['naphtha'])
        # dp.fill_carrier_data(casepath, value_or_data=2114, columns=['Import limit'], carriers=['methane_bio'])
        # dp.fill_carrier_data(casepath, value_or_data=990, columns=['Import limit'], carriers=['naphtha_bio'])
        # dp.fill_carrier_data(casepath, value_or_data=2359, columns=['Import limit'], carriers=['CO2_DAC'])
        # dp.fill_carrier_data(casepath, value_or_data=526, columns=['Import limit'], carriers=['ethanol'])
        # dp.fill_carrier_data(casepath, value_or_data=157, columns=['Import limit'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=338, columns=['Import limit'], carriers=['MPW'])

        # No export limit
        dp.fill_carrier_data(casepath, value_or_data=2000, columns=['Export limit'],
                             carriers=["nitrogen", "oxygen", "steam", "heatlowT", "crackergas"])

        # CO2 export
        dp.fill_carrier_data(casepath, value_or_data= 114, columns=['Export limit'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=-111.51, columns=['Export price'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=-1, columns=['Export emission factor'], carriers=['CO2'])

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=56, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=141, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=1830, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=618, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=1734, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=1473, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])

        # Constant import emission factor fossil feedstocks
        dp.fill_carrier_data(casepath, value_or_data=0.2, columns=['Import emission factor'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=3.23, columns=['Import emission factor'], carriers=['naphtha'])

        # Electricity price from file
        json_file_path = casepath / "Topology.json"
        with open(json_file_path, "r") as json_file:
            topology = json.load(json_file)

        for period in topology["investment_periods"]:
            for node_name in topology["nodes"]:
                # load electricity data
                el_load_path = Path(datapath) / 'import_data' / 'Electricity_data_MY.xlsx'
                el_importdata = pd.read_excel(el_load_path, sheet_name=period, header=0, nrows=8760)
                el_price = el_importdata.iloc[:, 0]
                el_emissionrate = el_importdata.iloc[:, 2]

                dp.fill_carrier_data(casepath, value_or_data=el_price, columns=['Import price'], carriers=['electricity'])
                dp.fill_carrier_data(casepath, value_or_data=el_emissionrate, columns=['Import emission factor'],
                                     carriers=['electricity'])

                # carbon tax
                file_path = Path(casepath) / period / "node_data" / node_name / 'CarbonCost.csv'
                data = pd.read_csv(file_path, delimiter=';')

                # Set the price to 300 and subsidy to 0 for all rows
                data['price'] = 150.31
                data['subsidy'] = 150.31

                # Save the modified CSV file
                data.to_csv(file_path, index=False, sep=';')


#Create data Chemelot cluster mid term
execute = 1

if execute == 1:
    # Specify the path to your input data
    casepath = basepath / "Case_studies" / "MY_Chemelot_bf_2040"
    datapath = basepath / "Input_data" / "250303_MY_Data_Chemelot_bf"

    firsttime = 0
    if firsttime:
        # Create template files
        dp.create_optimization_templates(casepath)
        dp.create_montecarlo_template_csv(casepath)

        # Change topology
        topology_file = casepath / "Topology.json"
        config_file = casepath / "ConfigModel.json"

        topology_template = {
            "nodes": ["Chemelot"],
            "carriers": ["electricity", "methane", "methane_bio", "naphtha", "naphtha_bio",
                         "CO2", "CO2_DAC", "CO2captured", "hydrogen", "nitrogen", "oxygen",
                         "ammonia", "ethylene", "propylene", "PE_olefin", "olefins",
                         "crackergas", "feedgas", "steam", "heatlowT", "HBfeed", "syngas", "syngas_r",
                         "methanol", "ethanol", "propane", "MPW"],
            "investment_periods": ["2040"],
            "start_date": "2022-01-01 00:00",
            "end_date": "2022-12-31 23:00",
            "resolution": "1h",
            "investment_period_length": 1,
        }

        # open and save topology and configuration files
        with open(topology_file, "w") as f:
            json.dump(topology_template, f, indent=4)
        # with open(config_file, "w") as f:
        # json.dump(configuration_template, f, indent=4)

        # Create folder structure
        dp.create_input_data_folder_template(casepath)

        # Node location
        file_path = casepath / 'NodeLocations.csv'
        data = pd.read_csv(file_path, delimiter=';')

        data['lon'] = 5.80
        data['lat'] = 50.97
        data['alt'] = 10

        # Save the modified CSV file
        data.to_csv(file_path, index=False, sep=';')

        # Read climate data and fill carried data
        dp.load_climate_data_from_api(casepath)

    read_techs = 1
    if read_techs or read_all:
        set_tecs = ["ElectricSMR_m", "WGS_m", "AEC", "HaberBosch",
                    "CrackerFurnace_Electric",
                    "ASU", "Boiler_Industrial_NG", "Boiler_El",
                    "RWGS", "MeOHsynthesis", "MTO", "EDH", "PDH", "MPW2methanol",
                    "DirectMeOHsynthesis", "CO2electrolysis",
                    "Storage_Ammonia", "Storage_CO2", "Storage_Ethylene",
                    "Storage_H2", "Storage_Battery", "Storage_Propylene",
                    "CO2toEmission", "feedgas_mixer", "naphtha_mixer", "PE_mixer", "CO2_mixer", "HBfeed_mixer", "syngas_mixer"]
        set_tecs_existing = {}

        json_file_path = casepath / "Topology.json"
        with open(json_file_path, "r") as json_file:
            topology = json.load(json_file)

        for period in topology["investment_periods"]:
            for node_name in topology["nodes"]:
                # Read the JSON technology file
                json_tec_file_path = (
                        casepath / period / "node_data" / node_name / "Technologies.json"
                )
                with open(json_tec_file_path, "r") as json_tec_file:
                    json_tec = json.load(json_tec_file)

                json_tec['new'] = set_tecs
                json_tec['existing'] = set_tecs_existing
                with open(json_tec_file_path, "w") as json_tec_file:
                    json.dump(json_tec, json_tec_file, indent=4)

                # Copy technology and network data into folder
                dp.copy_technology_data(casepath, datapath)
                
                # Add additional tecs
                set_tecfiles = ["MEA_large", "SteamReformer", "CrackerFurnace", "OlefinSeparation"]
                output_folder = (
                        casepath / period / "node_data" / node_name / "technology_data"
                )
                for tec in set_tecfiles:
                    other_tec_file_path = find_json_path(datapath, tec)
                    if other_tec_file_path:
                        shutil.copy(other_tec_file_path, output_folder)

                # specify tech size
                spec_tec_size = {"AEC": 2000, "RWGS": 287, "MeOHsynthesis": 632, "MTO": 1287, "SteamReformer": 1269,
                                 "ElectricSMR_m": 2114, "WGS_m": 389, "HaberBosch": 1609, "CrackerFurnace": 594,
                                 "CrackerFurnace_Electric": 990, "OlefinSeparation": 990, "ASU": 112,
                                 "Boiler_Industrial_NG": 2053, "Boiler_El": 2000, "EDH": 526, "PDH": 157,
                                 "MPW2methanol": 381, "DirectMeOHsynthesis": 301, "CO2electrolysis": 174}
                for tec in spec_tec_size.keys():
                    tech_data_file = find_json_path(output_folder, tec)
                    with open(tech_data_file, "r") as json_tecdata_file:
                        tech_data = json.load(json_tecdata_file)

                    tech_data["size_max"] = spec_tec_size[tec]
                    with open(tech_data_file, "w") as json_tecdata_file:
                        json.dump(tech_data, json_tecdata_file, indent=4)

    read_car = 0
    if read_car or read_all:
        # Fill carrier data
        dp.fill_carrier_data(casepath, value_or_data=0)

        # Demand data
        dp.fill_carrier_data(casepath, value_or_data=135, columns=['Demand'], carriers=['ammonia'])
        dp.fill_carrier_data(casepath, value_or_data=44, columns=['Demand'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=218, columns=['Demand'], carriers=['PE_olefin'])
        dp.fill_carrier_data(casepath, value_or_data=0, columns=['Demand'], carriers=['steam'])
        dp.fill_carrier_data(casepath, value_or_data=82, columns=['Demand'], carriers=['electricity'])
        dp.fill_carrier_data(casepath, value_or_data=7, columns=['Demand'], carriers=['feedgas'])

        # Import limits
        dp.fill_carrier_data(casepath, value_or_data=2000, columns=['Import limit'], carriers=['electricity'])
        dp.fill_carrier_data(casepath, value_or_data=5000, columns=['Import limit'],
                             carriers=['methane', 'naphtha', 'methane_bio', 'naphtha_bio', 'ethanol',
                                       'propane', 'MPW', 'CO2_DAC'])

        # dp.fill_carrier_data(casepath, value_or_data=2241, columns=['Import limit'], carriers=['methane'])
        # dp.fill_carrier_data(casepath, value_or_data=1093, columns=['Import limit'], carriers=['naphtha'])
        # dp.fill_carrier_data(casepath, value_or_data=2241, columns=['Import limit'], carriers=['methane_bio'])
        # dp.fill_carrier_data(casepath, value_or_data=1093, columns=['Import limit'], carriers=['naphtha_bio'])
        # dp.fill_carrier_data(casepath, value_or_data=2907, columns=['Import limit'], carriers=['CO2_DAC'])
        # dp.fill_carrier_data(casepath, value_or_data=559, columns=['Import limit'], carriers=['ethanol'])
        # dp.fill_carrier_data(casepath, value_or_data=173, columns=['Import limit'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=381, columns=['Import limit'], carriers=['MPW'])
        # dp.fill_carrier_data(casepath, value_or_data=338, columns=['Import limit'], carriers=['MPW'])

        # Import for demand
        # dp.fill_carrier_data(casepath, value_or_data=26, columns=['Import limit'], carriers=['PE_olefin'])
        # dp.fill_carrier_data(casepath, value_or_data=9, columns=['Import limit'], carriers=['ammonia'])
        # dp.fill_carrier_data(casepath, value_or_data=10000, columns=['Import price'], carriers=['PE_olefin', 'ammonia'])

        # No export limit
        dp.fill_carrier_data(casepath, value_or_data=2000, columns=['Export limit'],
                             carriers=["nitrogen", "oxygen", "steam", "heatlowT", "crackergas"])

        # CO2 export
        dp.fill_carrier_data(casepath, value_or_data=114, columns=['Export limit'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=-111.51, columns=['Export price'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=-1, columns=['Export emission factor'], carriers=['CO2'])

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=56, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=141, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=1830, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=475, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=1734, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=1473, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])

        # Constant import emission factor fossil feedstocks
        dp.fill_carrier_data(casepath, value_or_data=0.2, columns=['Import emission factor'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=3.23, columns=['Import emission factor'], carriers=['naphtha'])

        # Electricity price from file
        json_file_path = casepath / "Topology.json"
        with open(json_file_path, "r") as json_file:
            topology = json.load(json_file)

        for period in topology["investment_periods"]:
            for node_name in topology["nodes"]:
                # load electricity data
                el_load_path = Path(datapath) / 'import_data' / 'Electricity_data_MY.xlsx'
                el_importdata = pd.read_excel(el_load_path, sheet_name=period, header=0, nrows=8760)
                el_price = el_importdata.iloc[:, 0]
                el_emissionrate = el_importdata.iloc[:, 2]

                dp.fill_carrier_data(casepath, value_or_data=el_price, columns=['Import price'],
                                     carriers=['electricity'])
                dp.fill_carrier_data(casepath, value_or_data=el_emissionrate,
                                     columns=['Import emission factor'],
                                     carriers=['electricity'])

                #carbon tax
                file_path = Path(casepath) / period / "node_data" / node_name / 'CarbonCost.csv'
                data = pd.read_csv(file_path, delimiter=';')

                # Set the price to 300 and subsidy to 0 for all rows
                data['price'] = 150.31
                data['subsidy'] = 150.31

                # Save the modified CSV file
                data.to_csv(file_path, index=False, sep=';')




#Create data Chemelot cluster long term
execute = 1

if execute == 1:
    # Specify the path to your input data
    casepath = basepath / "Case_studies" / "MY_Chemelot_bf_2050"
    datapath = basepath / "Input_data" / "250303_MY_Data_Chemelot_bf"

    firsttime = 0
    if firsttime:
        # Create template files
        dp.create_optimization_templates(casepath)
        dp.create_montecarlo_template_csv(casepath)

        # Change topology
        topology_file = casepath / "Topology.json"
        config_file = casepath / "ConfigModel.json"

        topology_template = {
            "nodes": ["Chemelot"],
            "carriers": ["electricity", "methane", "methane_bio", "naphtha", "naphtha_bio",
                         "CO2", "CO2_DAC", "CO2captured", "hydrogen", "nitrogen", "oxygen",
                         "ammonia", "ethylene", "propylene", "PE_olefin", "olefins",
                         "crackergas", "feedgas", "steam", "heatlowT", "HBfeed", "syngas", "syngas_r",
                         "methanol", "ethanol", "propane", "MPW"],
            "investment_periods": ["2050"],
            "start_date": "2022-01-01 00:00",
            "end_date": "2022-12-31 23:00",
            "resolution": "1h",
            "investment_period_length": 1,
        }

        # open and save topology and configuration files
        with open(topology_file, "w") as f:
            json.dump(topology_template, f, indent=4)
        # with open(config_file, "w") as f:
        # json.dump(configuration_template, f, indent=4)

        # Create folder structure
        dp.create_input_data_folder_template(casepath)

        # Node location
        file_path = casepath / 'NodeLocations.csv'
        data = pd.read_csv(file_path, delimiter=';')

        data['lon'] = 5.80
        data['lat'] = 50.97
        data['alt'] = 10

        # Save the modified CSV file
        data.to_csv(file_path, index=False, sep=';')

        # Read climate data and fill carried data
        dp.load_climate_data_from_api(casepath)

    read_techs = 1
    if read_techs or read_all:
        set_tecs = ["ElectricSMR_m", "WGS_m", "AEC", "HaberBosch",
                    "CrackerFurnace_Electric",
                    "ASU", "Boiler_Industrial_NG", "Boiler_El",
                    "RWGS", "MeOHsynthesis", "MTO", "EDH", "PDH", "MPW2methanol",
                    "DirectMeOHsynthesis", "CO2electrolysis",
                    "Storage_Ammonia", "Storage_CO2", "Storage_Ethylene",
                    "Storage_H2", "Storage_Battery", "Storage_Propylene",
                    "CO2toEmission", "feedgas_mixer", "naphtha_mixer", "PE_mixer", "CO2_mixer", "HBfeed_mixer", "syngas_mixer"]
        set_tecs_existing = {}

        json_file_path = casepath / "Topology.json"
        with open(json_file_path, "r") as json_file:
            topology = json.load(json_file)

        for period in topology["investment_periods"]:
            for node_name in topology["nodes"]:
                # Read the JSON technology file
                json_tec_file_path = (
                        casepath / period / "node_data" / node_name / "Technologies.json"
                )
                with open(json_tec_file_path, "r") as json_tec_file:
                    json_tec = json.load(json_tec_file)

                json_tec['new'] = set_tecs
                json_tec['existing'] = set_tecs_existing

                with open(json_tec_file_path, "w") as json_tec_file:
                    json.dump(json_tec, json_tec_file, indent=4)

                # Copy technology and network data into folder
                dp.copy_technology_data(casepath, datapath)

                # Add additional tecs
                set_tecfiles = ["MEA_large", "SteamReformer", "CrackerFurnace", "OlefinSeparation"]
                output_folder = (
                        casepath / period / "node_data" / node_name / "technology_data"
                )
                for tec in set_tecfiles:
                    other_tec_file_path = find_json_path(datapath, tec)
                    if other_tec_file_path:
                        shutil.copy(other_tec_file_path, output_folder)

                # specify tech size
                spec_tec_size = {"AEC": 2500, "RWGS": 327, "MeOHsynthesis": 666, "MTO": 1408, "SteamReformer": 1269,
                                 "ElectricSMR_m": 2114, "WGS_m": 389, "HaberBosch": 1609, "CrackerFurnace": 594,
                                 "CrackerFurnace_Electric": 990, "OlefinSeparation": 990, "ASU": 112,
                                 "Boiler_Industrial_NG": 2053, "Boiler_El": 2500, "EDH": 526, "PDH": 157,
                                 "MPW2methanol": 424, "DirectMeOHsynthesis": 344, "CO2electrolysis": 218}
                for tec in spec_tec_size.keys():
                    tech_data_file = find_json_path(output_folder, tec)
                    with open(tech_data_file, "r") as json_tecdata_file:
                        tech_data = json.load(json_tecdata_file)

                    tech_data["size_max"] = spec_tec_size[tec]
                    with open(tech_data_file, "w") as json_tecdata_file:
                        json.dump(tech_data, json_tecdata_file, indent=4)

    read_car = 0
    if read_car or read_all:
        # Fill carrier data
        dp.fill_carrier_data(casepath, value_or_data=0)

        # Demand data
        dp.fill_carrier_data(casepath, value_or_data=135, columns=['Demand'], carriers=['ammonia'])
        dp.fill_carrier_data(casepath, value_or_data=44, columns=['Demand'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=218, columns=['Demand'], carriers=['PE_olefin'])
        dp.fill_carrier_data(casepath, value_or_data=0, columns=['Demand'], carriers=['steam'])
        dp.fill_carrier_data(casepath, value_or_data=82, columns=['Demand'], carriers=['electricity'])
        dp.fill_carrier_data(casepath, value_or_data=7, columns=['Demand'], carriers=['feedgas'])

        # Import limits
        dp.fill_carrier_data(casepath, value_or_data=2500, columns=['Import limit'], carriers=['electricity'])
        dp.fill_carrier_data(casepath, value_or_data=5000, columns=['Import limit'],
                             carriers=['methane', 'naphtha', 'methane_bio', 'naphtha_bio', 'ethanol',
                                       'propane', 'MPW', 'CO2_DAC'])

        # dp.fill_carrier_data(casepath, value_or_data=2382, columns=['Import limit'], carriers=['methane'])
        # dp.fill_carrier_data(casepath, value_or_data=1210, columns=['Import limit'], carriers=['naphtha'])
        # dp.fill_carrier_data(casepath, value_or_data=2382, columns=['Import limit'], carriers=['methane_bio'])
        # dp.fill_carrier_data(casepath, value_or_data=1210, columns=['Import limit'], carriers=['naphtha_bio'])
        # dp.fill_carrier_data(casepath, value_or_data=3218, columns=['Import limit'], carriers=['CO2_DAC'])
        # dp.fill_carrier_data(casepath, value_or_data=619, columns=['Import limit'], carriers=['ethanol'])
        # dp.fill_carrier_data(casepath, value_or_data=192, columns=['Import limit'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=424, columns=['Import limit'], carriers=['MPW'])
        # dp.fill_carrier_data(casepath, value_or_data=338, columns=['Import limit'], carriers=['MPW'])

        # Import for demand
        # dp.fill_carrier_data(casepath, value_or_data=26, columns=['Import limit'], carriers=['PE_olefin'])
        # dp.fill_carrier_data(casepath, value_or_data=9, columns=['Import limit'], carriers=['ammonia'])
        # dp.fill_carrier_data(casepath, value_or_data=10000, columns=['Import price'], carriers=['PE_olefin', 'ammonia'])

        # No export limit
        dp.fill_carrier_data(casepath, value_or_data=2000, columns=['Export limit'],
                             carriers=["nitrogen", "oxygen", "steam", "heatlowT", "crackergas"])

        # CO2 export
        dp.fill_carrier_data(casepath, value_or_data=148, columns=['Export limit'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=-111.51, columns=['Export price'], carriers=['CO2'])
        dp.fill_carrier_data(casepath, value_or_data=-1, columns=['Export emission factor'], carriers=['CO2'])

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=59, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=148, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=1830, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=355, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=1734, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=1473, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])

        # Constant import emission factor fossil feedstocks
        dp.fill_carrier_data(casepath, value_or_data=0.2, columns=['Import emission factor'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=3.23, columns=['Import emission factor'], carriers=['naphtha'])

        # Electricity price from file:
        json_file_path = casepath / "Topology.json"
        with open(json_file_path, "r") as json_file:
            topology = json.load(json_file)

        for period in topology["investment_periods"]:
            for node_name in topology["nodes"]:
                # load electricity data
                el_load_path = Path(datapath) / 'import_data' / 'Electricity_data_MY.xlsx'
                el_importdata = pd.read_excel(el_load_path, sheet_name=period, header=0, nrows=8760)
                el_price = el_importdata.iloc[:, 0]
                el_emissionrate = el_importdata.iloc[:, 2]

                dp.fill_carrier_data(casepath, value_or_data=el_price, columns=['Import price'],
                                     carriers=['electricity'])
                dp.fill_carrier_data(casepath, value_or_data=el_emissionrate,
                                     columns=['Import emission factor'],
                                     carriers=['electricity'])

                # carbon tax
                file_path = Path(casepath) / period / "node_data" / node_name / 'CarbonCost.csv'
                data = pd.read_csv(file_path, delimiter=';')

                # Set the price to 400 and subsidy to 0 for all rows
                data['price'] = 150.31
                data['subsidy'] = 150.31

                # Save the modified CSV file
                data.to_csv(file_path, index=False, sep=';')

