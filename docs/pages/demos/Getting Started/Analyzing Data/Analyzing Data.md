---
title: Analyzing Data
has_children: false
parent: Getting started
grand_parent: Demos
nav_order: 4
---

# Step by step

## Overview


### Viewing the results

All of the results obtained from running the application are written in the result folder located in `<repo>/code/results`. 

A subfolder named `<(year)_monthly_results>` has all of the information of a lattice Ran that specific year. In this case, it has the data of 2021 mice lattices.
This folder contains the different files that are being saved, specific of the mice lattice the user is looking into. It is divided into two parts:

+ The excel workbook, always saved as `<data_results>`. 

+ A `<plots_per_month>` folder which contains all of the plots created by the user with the application. 

Both the excel workbook and the plots_per_month folder carry information the user created with the application and will be explained next.

### Understanding the results

<img src='output_folder.png'>

The file `summary.png` shows pCa, length, force per cross-sectional area (stress), and thick and thin filamnt properties plotted against time..

<img src='summary.png' width="50%">

The underlying data are stored in `results.txt`

<img src='results.png' width="100%">

## How this worked

This demonstration simulated a half-sarcomere that was held isometric and activated in a solution with a pCa of 4.5.

The simulation was controlled by a batch file (shown below) that was passed to FiberPy.

The first few lines, labelled `FiberSim_batch` tell FiberPy where to find the `FiberCpp.exe` file which is the core model.

The rest of the file defines a single `job`. FiberPy passes the job to the core model which runs the simulation and saves the results.

Each job consists of:

+ a model file - which defines the properties of the half-sarcomere including the structure, and the biophysical parameters that define the kinetic schemes for the thick and thin filaments
+ a protocol file - the pCa value and information about whether the system is in length control or force-control mode
+ an options file - which can be used to set additional criteria
+ a results file - which stores information about the simulation

````
{
    "FiberSim_batch": {
        "FiberCpp_exe":
        {
            "relative_to": "this_file",
            "exe_file": "../../../bin/FiberCpp.exe"
        },
        "job":[
            {
                "relative_to": "this_file",
                "model_file": "sim_input/model.json",
                "options_file": "sim_input/options.json",
                "protocol_file": "sim_input/pCa45_protocol.txt",
                "results_file": "sim_output/results.txt",
                "output_handler_file": "sim_input/output_handler.json"
            }
        ]
    }
}
````

The last entry in the job is optional and defines an output-handler. In this example `output_handler.json` was

````
{
    "templated_images":
    [
        {
            "relative_to": "this_file",
            "template_file_string": "../template/template_summary.json",
            "output_file_string": "../sim_output/summary.png"
        }
    ]
}
````

This instructed FiberPy to:
+ take the simulation results
+ create a figure using the framework described in the `template_summary.json` file
+ save the data to `../sim_output/summary.png`

