# Explainable Sustainability Assessment and Optimization Framework for Toluene HDA Process

## Overview

This project presents a data-driven and explainable framework for the sustainability assessment and optimization of a full toluene hydrodealkylation (HDA) flowsheet, integrating global sensitivity analysis (GSA), surrogate modeling, and Bayesian optimization (BO) within a composite Life-Cycle Sustainability Index (LCSI) framework.

The methodology is designed to address the intrinsically coupled and multi-objective nature of process-level sustainability, simultaneously considering economic, environmental, social, and technological performance metrics while maintaining high computational efficiency.

<img width="1193" height="712" alt="Framework overview" src="https://github.com/user-attachments/assets/ae1d8f72-0d16-40ca-917c-ac6057f7beee" />

Within this broader framework, the current folder provides a self-contained supertargeting and pinch-analysis demo based on an Excel stream sheet. It is intended as a compact, reproducible example for heat-integration analysis within the overall sustainability workflow.



## HEN Supertargeting Demo

This submodule focuses on the heat-integration part of the broader workflow. The demo covers:

- reading thermal streams from Excel
- converting the rows into `ThermalStream` objects
- running automatic supertargeting at a selected `\Delta T_{\min}`
- reviewing summary metrics and area intervals
- comparing composite curve and balanced composite curve results
- displaying the final rendered figures used in the workflow

## Included Files

- `Input Sheet of 20.xlsx`: local demo data source
- `notebook_sources/supertargeting_demo.py`: notebook source of truth
- `generate_notebooks.py`: thin dispatcher that writes the notebook
- `supertargeting_demo.ipynb`: generated Jupyter notebook
- `images/`: rendered figures used by the notebook demo

## Notebook Outputs

The notebook is set up to show:

- stream input table
- summary metrics table
- consolidated total-results table
- interval and area table
- notebook-rendered CC and BCC plots
- pre-rendered final figures from the workflow

## Usage

Open `supertargeting_demo.ipynb` directly in Jupyter and run the cells from top to bottom.


## Acknowledgement

The authors also gratefully acknowledge A/Prof. Sachin V. Jangam for for his teaching in `CN4205 Pinch Analysis and Process Integration`, from which this work benefited.

## Contact

Questions, feedback, or collaboration ideas are welcome. Please contact: `ziyun_zhang@u.nus.edu`
