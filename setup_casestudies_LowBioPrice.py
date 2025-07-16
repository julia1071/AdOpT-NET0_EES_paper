import json
import shutil
from pathlib import Path
import adopt_net0.data_preprocessing as dp
from adopt_net0.data_preprocessing.data_loading import find_json_path
from adopt_net0.modelhub import ModelHub
from adopt_net0.result_management.read_results import add_values_to_summary
import pandas as pd

#global functions
execute_greenfield = 1
execute_brownfield = 1
read_all_greenfield = 1
read_all_brownfield = 1


if execute_greenfield:
    #Create data Chemelot cluster short term
    execute = 1

    if execute == 1:
        # Specify the path to your input data
        casepath = Path("Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_gf_2030_OptBIO")

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=56, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=113, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=1464, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=355, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=1387, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=1178, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])


    #Create data Chemelot cluster mid term
    execute = 1

    if execute == 1:
        # Specify the path to your input data
        casepath = Path("Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_gf_2040_OptBIO")

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=56, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=85, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=1098, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=275, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=1040, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=884, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])


    #Create data Chemelot cluster long term
    execute = 1

    if execute == 1:
        # Specify the path to your input data
        casepath = Path("Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_gf_2050_OptBIO")

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=59, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=59, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=200, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=694, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=589, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])


if execute_brownfield:
    #Create data Chemelot cluster short term
    execute = 1

    if execute == 1:
        # Specify the path to your input data
        casepath = Path("Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_bf_2030_OptBIO")

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=56, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=113, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=1464, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=355, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=1387, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=1178, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])


    #Create data Chemelot cluster mid term
    execute = 1

    if execute == 1:
        # Specify the path to your input data
        casepath = Path("Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_bf_2040_OptBIO")

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=56, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=85, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=1098, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=275, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=1040, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=884, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])


    #Create data Chemelot cluster long term
    execute = 1

    if execute == 1:
        # Specify the path to your input data
        casepath = Path("Z:/AdOpt_NET0/AdOpt_casestudies/MY/MY_Chemelot_bf_2050_OptBIO")

        # Constant import prices
        dp.fill_carrier_data(casepath, value_or_data=59, columns=['Import price'], carriers=['methane'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha'])
        dp.fill_carrier_data(casepath, value_or_data=59, columns=['Import price'], carriers=['methane_bio'])
        dp.fill_carrier_data(casepath, value_or_data=732, columns=['Import price'], carriers=['naphtha_bio'])
        dp.fill_carrier_data(casepath, value_or_data=200, columns=['Import price'], carriers=['CO2_DAC'])
        dp.fill_carrier_data(casepath, value_or_data=694, columns=['Import price'], carriers=['ethanol'])
        dp.fill_carrier_data(casepath, value_or_data=589, columns=['Import price'], carriers=['propane'])
        dp.fill_carrier_data(casepath, value_or_data=780, columns=['Import price'], carriers=['MPW'])

