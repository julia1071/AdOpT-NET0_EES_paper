import pytest
import pandas as pd
import os

import adopt_net0.data_preprocessing as dp
from adopt_net0.data_management.utilities import calculate_dni
from adopt_net0.data_management.utilities import check_input_data_consistency
from tests.utilities import (
    select_random_list_from_list,
    load_json,
    save_json,
    get_topology_data,
)


def test_create_input_data_folder(request):
    """
    Tests standard behavior of
    - create_optimization_templates
    - initialize_configuration_templates
    - initialize_topology_templates
    - create_input_data_folder_template
    - create_empty_network_matrix
    """
    data_folder_path = request.config.data_folder_path
    dp.create_optimization_templates(data_folder_path)
    dp.create_input_data_folder_template(data_folder_path)


def test_data_climate_data_loading(request):
    """
    Tests standard behavior of load_climate_data_from_api
    - Tests if df is not empty
    - Tests if climate data is the same in investment periods
    """
    case_study_folder_path = request.config.case_study_folder_path

    # Write it to file
    dp.load_climate_data_from_api(case_study_folder_path)

    # Get periods and nodes:
    investment_periods, nodes, carriers = get_topology_data(case_study_folder_path)

    # Verify that climate data is not empty
    climate_data = {}
    for period in investment_periods:
        climate_data[period] = {}
        for node in nodes:
            climate_data[period][node] = pd.read_csv(
                case_study_folder_path
                / period
                / "node_data"
                / node
                / "ClimateData.csv",
                sep=";",
                index_col=0,
            )
            assert not climate_data[period][node].empty

            # calculate dni and check if its ok
            node_locations = pd.read_csv(
                case_study_folder_path / "NodeLocations.csv", sep=";", index_col=0
            )
            lon = node_locations.loc[node, "lon"]
            lat = node_locations.loc[node, "lat"]

            climate_data_check = climate_data[period][node]
            climate_data_check["dni_correct"] = climate_data_check["dni"]
            climate_data_check = climate_data_check.drop(columns=["dni"])
            climate_data_check["dni_check"] = calculate_dni(
                climate_data_check, lon, lat
            )


def test_data_fill_carrier_data(request):
    """
    Tests standard behavior of fill_carrier_data
    - Tests if df is indeed filled
    """
    case_study_folder_path = request.config.case_study_folder_path

    # Write to files (should fill some random columns with 1 for all nodes and investment periods
    carriers_to_fill = ["electricity"]
    fill_options = [
        "Demand",
        "Import limit",
        "Export limit",
        "Import price",
        "Export price",
        "Import emission factor",
        "Export emission factor",
        "Generic production",
    ]
    series_to_fill = select_random_list_from_list(fill_options)

    dp.fill_carrier_data(
        case_study_folder_path, 1, columns=series_to_fill, carriers=carriers_to_fill
    )

    # Check if it is filled indeed
    # Get periods and nodes:
    investment_periods, nodes, carriers = get_topology_data(case_study_folder_path)

    # Verify that climate data is not empty
    for period in investment_periods:
        for node in nodes:
            for carrier in carriers:
                carrier_data = pd.read_csv(
                    case_study_folder_path
                    / period
                    / "node_data"
                    / node
                    / "carrier_data"
                    / (carrier + ".csv"),
                    sep=";",
                )
                for col in carrier_data.columns:
                    if (carrier in carriers_to_fill) and (col in series_to_fill):
                        assert (carrier_data[col] == 1).all()


def test_copy_technology_data(request):
    """
    Tests standard behavior of fill_carrier_data
    - Tests if df is indeed filled
    """
    case_study_folder_path = request.config.case_study_folder_path
    technology_data_folder_path = request.config.technology_data_folder_path

    investment_periods, nodes, carriers = get_topology_data(case_study_folder_path)
    nodes_to_add_to = select_random_list_from_list(nodes)
    periods_to_add_to = select_random_list_from_list(investment_periods)

    # Create technologies
    for period in periods_to_add_to:
        for node in nodes_to_add_to:
            path = (
                case_study_folder_path
                / period
                / "node_data"
                / node
                / "Technologies.json"
            )
            technologies = load_json(path)
            technologies["existing"] = {"TestTec_Conv1": 5}
            technologies["new"] = ["TestTec_Conv2", "TestTec_GasTurbine_simple_CCS"]
            save_json(technologies, path)

    # Copy to folder
    dp.copy_technology_data(case_study_folder_path, technology_data_folder_path)

    # Check it jsons are there
    check_input_data_consistency(case_study_folder_path)


def test_copy_network_data(request):
    """
    Tests standard behavior of fill_carrier_data
    - Tests if df is indeed filled
    """
    case_study_folder_path = request.config.case_study_folder_path
    network_data_folder_path = request.config.network_data_folder_path

    investment_periods, nodes, carriers = get_topology_data(case_study_folder_path)
    periods_to_add_to = select_random_list_from_list(investment_periods)

    # Create networks
    for period in periods_to_add_to:
        path = case_study_folder_path / period / "Networks.json"
        networks = load_json(path)
        networks["new"] = ["TestNetworkSimple"]
        save_json(networks, path)

        os.makedirs(
            case_study_folder_path
            / period
            / "network_topology"
            / "new"
            / "TestNetworkSimple"
        )

        connection = pd.read_csv(
            case_study_folder_path
            / period
            / "network_topology"
            / "new"
            / "connection.csv",
            sep=";",
            index_col=0,
        )
        connection.to_csv(
            case_study_folder_path
            / period
            / "network_topology"
            / "new"
            / "TestNetworkSimple"
            / "connection.csv",
            sep=";",
        )
        os.remove(
            case_study_folder_path
            / period
            / "network_topology"
            / "new"
            / "connection.csv"
        )

        distance = pd.read_csv(
            case_study_folder_path
            / period
            / "network_topology"
            / "new"
            / "distance.csv",
            sep=";",
            index_col=0,
        )
        distance.loc["city", "rural"] = 50
        distance.loc["rural", "city"] = 50
        distance.to_csv(
            case_study_folder_path
            / period
            / "network_topology"
            / "new"
            / "TestNetworkSimple"
            / "distance.csv",
            sep=";",
        )
        os.remove(
            case_study_folder_path
            / period
            / "network_topology"
            / "new"
            / "distance.csv"
        )

    # Copy to folder
    dp.copy_network_data(case_study_folder_path, network_data_folder_path)

    # Check it jsons are there
    check_input_data_consistency(case_study_folder_path)
