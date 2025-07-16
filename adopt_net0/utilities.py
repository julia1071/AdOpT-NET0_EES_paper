import json
from pathlib import Path

import h5py
import pandas as pd
from pyomo.environ import SolverFactory

# from adopt_net0.result_management.read_results import extract_datasets_from_h5group


def get_gurobi_parameters(solveroptions: dict):
    """
    Initiates the gurobi solver and defines solver parameters

    :param dict solveroptions: dict with solver parameters
    :return: Gurobi Solver
    """
    solver = SolverFactory(solveroptions["solver"]["value"], solver_io="python")
    solver.options["TimeLimit"] = solveroptions["timelim"]["value"] * 3600
    solver.options["MIPGap"] = solveroptions["mipgap"]["value"]
    solver.options["MIPFocus"] = solveroptions["mipfocus"]["value"]
    solver.options["Threads"] = solveroptions["threads"]["value"]
    solver.options["NodefileStart"] = solveroptions["nodefilestart"]["value"]
    solver.options["Method"] = solveroptions["method"]["value"]
    solver.options["Heuristics"] = solveroptions["heuristics"]["value"]
    solver.options["Presolve"] = solveroptions["presolve"]["value"]
    solver.options["BranchDir"] = solveroptions["branchdir"]["value"]
    solver.options["LPWarmStart"] = solveroptions["lpwarmstart"]["value"]
    solver.options["IntFeasTol"] = solveroptions["intfeastol"]["value"]
    solver.options["FeasibilityTol"] = solveroptions["feastol"]["value"]
    solver.options["Cuts"] = solveroptions["cuts"]["value"]
    solver.options["NumericFocus"] = solveroptions["numericfocus"]["value"]
    # solver.options["SubMIPNodes"] = solveroptions["SubMIPNodes"]["value"]
    # solver.options["ScaleFlag"] = solveroptions["ScaleFlag"]["value"]

    return solver


def get_glpk_parameters(solveroptions: dict):
    """
    Initiates the glpk solver and defines solver parameters

    :param dict solveroptions: dict with solver parameters
    :return: Gurobi Solver
    """
    solver = SolverFactory("glpk")

    return solver


def get_set_t(config: dict, model_block):
    """
    Returns the correct set_t for different clustering options

    :param dict config: config dict
    :param model_block: pyomo block holding set_t_full and set_t_clustered
    :return: set_t
    """
    if config["optimization"]["typicaldays"]["N"]["value"] == 0:
        return model_block.set_t_full
    elif config["optimization"]["typicaldays"]["method"]["value"] == 1:
        return model_block.set_t_clustered
    elif config["optimization"]["typicaldays"]["method"]["value"] == 2:
        return model_block.set_t_full


def get_hour_factors(config: dict, data, period: str) -> list:
    """
    Returns the correct hour factors to use for global balances

    :param dict config: config dict
    :param data: DataHandle
    :return: hour factors
    """
    if config["optimization"]["typicaldays"]["N"]["value"] == 0:
        return [1] * len(data.topology["time_index"]["full"])
    elif config["optimization"]["typicaldays"]["method"]["value"] == 1:
        return data.k_means_specs[period]["factors"]
    elif config["optimization"]["typicaldays"]["method"]["value"] == 2:
        return [1] * len(data.topology["time_index"]["full"])


def get_nr_timesteps_averaged(config: dict) -> int:
    """
    Returns the correct number of timesteps averaged

    :param dict config: config dict
    :return: nr_timesteps_averaged
    """
    if config["optimization"]["timestaging"]["value"] != 0:
        nr_timesteps_averaged = config["optimization"]["timestaging"]["value"]
    else:
        nr_timesteps_averaged = 1

    return nr_timesteps_averaged


def fix_technology_sizes_zero(m, all_tecs, node, period):
    for tec in all_tecs.keys():
        if tec in m.model[m.info_solving_algorithms["aggregation_model"]].periods[period].node_blocks[node].tech_blocks_active:
            if all_tecs[tec]:
                m.model[m.info_solving_algorithms["aggregation_model"]].periods[period].node_blocks[node].tech_blocks_active[tec].var_size.fix(0)
            else:
                m.model[m.info_solving_algorithms["aggregation_model"]].periods[period].node_blocks[node].tech_blocks_active[tec].var_size.unfix()

    return m


def fix_installed_capacities(m, scenario, prev_scenario, node, exceptions):
    """
    Set the capacity of a technology to a minimum capacity for the next brownfield simulation
    """
    model = m[scenario].model[m[scenario].info_solving_algorithms["aggregation_model"]].periods[scenario]
    prev_model = m[prev_scenario].model[m[scenario].info_solving_algorithms["aggregation_model"]].periods[prev_scenario]

    b_node = model.node_blocks[node]
    b_node_prev = prev_model.node_blocks[node]

    for tec_name in b_node_prev.set_technologies:
        if tec_name not in exceptions:
            b_tec = b_node.tech_blocks_active[tec_name]
            prev_tec_size = b_node_prev.tech_blocks_active[tec_name].var_size.value

            b_tec.var_size.setlb(prev_tec_size)

    return m


def installed_capacities_existing(m, scenario, prev_scenario, node, casepath, delayed=0):
    """
    Set the capacity of a technology to a minimum capacity for the next brownfield simulation
    """
    if not delayed:
        prev_model = m[prev_scenario].model[m[prev_scenario].info_solving_algorithms["aggregation_model"]].periods[prev_scenario]

        b_node_prev = prev_model.node_blocks[node]

    size_tecs_existing = {}

    if delayed:
        size_tecs_existing = {"SteamReformer": 1078, "HaberBosch": 813, "CrackerFurnace": 499, "OlefinSeparation": 499}
    else:
        # Loop through all technologies
        for tec_name in b_node_prev.set_technologies:
            if tec_name.endswith("_existing"):
                # Standalone existing technology (no new counterpart)
                base_tec_name = tec_name.replace("_existing", "")
                if base_tec_name not in b_node_prev.set_technologies:
                    prev_existing_size = b_node_prev.tech_blocks_active[tec_name].var_size.value or 0
                    if prev_existing_size > 0:
                        size_tecs_existing[base_tec_name] = prev_existing_size
                continue  # Skip processing it as a "new" technology

            # New technology case
            prev_tec_size = b_node_prev.tech_blocks_active[tec_name].var_size.value or 0

            existing_tec_name = tec_name + "_existing"
            prev_existing_size = 0

            if existing_tec_name in b_node_prev.set_technologies:
                prev_existing_size = b_node_prev.tech_blocks_active[existing_tec_name].var_size.value or 0

            if prev_tec_size + prev_existing_size > 0:
                size_tecs_existing[tec_name] = prev_tec_size + prev_existing_size


    # Read the JSON technology file
    casepath = Path(casepath)
    json_tec_file_path = (
            casepath / scenario / "node_data" / node / "Technologies.json"
    )
    with open(json_tec_file_path, "r") as json_tec_file:
        json_tec = json.load(json_tec_file)

    json_tec['existing'] = size_tecs_existing
    with open(json_tec_file_path, "w") as json_tec_file:
        json.dump(json_tec, json_tec_file, indent=4)


def installed_capacities_existing_from_file(scenario, prev_scenario, node, casepath, h5_path_prev):
    """
    Set the capacity of a technology to a minimum capacity for the next brownfield simulation
    """
    size_tecs_existing = {}

    with h5py.File(h5_path_prev, "r") as hdf_file:
        nodedata = extract_datasets_from_h5group2(hdf_file["design/nodes"])
        df_nodedata = pd.DataFrame(nodedata)

        for tec_name in df_nodedata.columns.levels[2]:
            if (prev_scenario, node, tec_name, 'size') in df_nodedata.columns:
                prev_tec_size = df_nodedata[(prev_scenario, node, tec_name, 'size')].iloc[0]

                if prev_tec_size > 0:
                    if 'existing' in tec_name:
                        tec_name = tec_name[:-9]
                    size_tecs_existing[tec_name] = prev_tec_size

    # Read the JSON technology file
    casepath = Path(casepath)
    json_tec_file_path = (
            casepath / scenario / "node_data" / node / "Technologies.json"
    )
    with open(json_tec_file_path, "r") as json_tec_file:
        json_tec = json.load(json_tec_file)

    json_tec['existing'] = size_tecs_existing
    with open(json_tec_file_path, "w") as json_tec_file:
        json.dump(json_tec, json_tec_file, indent=4)


def extract_datasets_from_h5group2(group, prefix: tuple = ()) -> dict:
    """
    Extracts datasets from a group within a h5 file

    Gets all datasets from a group of a h5 file and writes it to a multi-index
    dataframe using a recursive function

    :param group: froup of h5 file
    :param tuple prefix: required to search through the structure of the h5 tree if there are multiple subgroups in the
     group you specified, empty by default meaning it starts searching from the group specified.
    :return: dataframe containing all datasets in group
    :rtype: pd.DataFrame
    """
    data = {}
    for key, value in group.items():
        if isinstance(value, h5py.Group):
            data.update(extract_datasets_from_h5group2(value, prefix + (key,)))
        elif isinstance(value, h5py.Dataset):
            if value.shape == ():
                data[prefix + (key,)] = [value[()]]
            else:
                data[prefix + (key,)] = value[:]

    return data