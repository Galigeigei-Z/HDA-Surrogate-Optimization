# Explainable Sustainability Assessment and Optimization Framework for Toluene HDA Process

## Overview

This project presents a data-driven and explainable framework for the sustainability assessment and optimization of a full toluene hydrodealkylation (HDA) flowsheet, integrating global sensitivity analysis (GSA), surrogate modeling, and Bayesian optimization (BO) within a composite Life-Cycle Sustainability Index (LCSI) framework.

The methodology is designed to address the intrinsically coupled and multi-objective nature of process-level sustainability, simultaneously considering economic, environmental, social, and technological performance metrics while maintaining high computational efficiency.

<img width="1193" height="712" alt="Framework overview" src="https://github.com/user-attachments/assets/ae1d8f72-0d16-40ca-917c-ac6057f7beee" />

Within this broader framework, the current folder provides a self-contained supertargeting and pinch-analysis demo based on an Excel stream sheet. It is intended as a compact, reproducible example for heat-integration analysis within the overall sustainability workflow.

## Methodology

A high-fidelity HDA flowsheet producing 100 kt/yr of ultra-high-purity benzene (99.99 wt%) was developed in Aspen HYSYS and systematically analyzed.

Key methodological components include:

- Composite LCSI formulation integrating 7 sustainability indicators
  economic: TOC, ACC
  environmental: CO2, SO2, NOx
  social: process risk index
  technological: exergy efficiency
- Perturbation-based global sensitivity analysis to reduce the decision space from 14 variables to a compact set of critical operating parameters using only 106 converged simulations
- ANN-based surrogate models trained on the reduced parameter space with high predictive accuracy
- Explainable AI (SHAP) analysis to identify dominant sustainability drivers and nonlinear trade-offs
- GSA-enhanced ANN-assisted Bayesian optimization for efficient multi-objective decision-making

## Key Results

- ANN surrogate models achieved MAPE below 1.0% for all sustainability indicators, with R2 exceeding 0.995 for most outputs
- SHAP analysis revealed that purge ratio, reactor operation mode, and feed temperature dominate sustainability performance, while separation-related parameters play secondary roles
- Surrogate-assisted Bayesian optimization improved the composite LCSI by 17.1% relative to nominal operation
- Compared with simulation-based optimization, the proposed framework achieved about 99% reduction in computational cost while maintaining solution deviations within a few percent

## Robustness Analysis

The optimized operating conditions were further evaluated under time-varying electricity prices and grid emission factors in Singapore from 2005 to 2024. Results show that the improved sustainability performance remains relatively stable across evolving energy-system conditions, indicating that the LCSI framework can reconcile competing economic and environmental drivers.

## Significance

This work shows that explainable, surrogate-assisted sustainability optimization can provide mechanistically consistent insights beyond black-box optimization while dramatically reducing computational burden. The proposed GSA-ANN-BO framework offers a scalable and reliable tool for:

- process-level sustainability assessment
- energy- and emission-aware operation optimization
- decision support in chemical reaction engineering systems

## Future Work

Future extensions will focus on:

- structural process modifications and process intensification
- energy-intensive separation units using mechanical and fluid vapor recompression (MVR/FVR)
- exploiting peak-off-peak electricity price and emission variability
- evaluating thermal energy storage integrated FVR configurations

## Supertargeting Demo

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

To regenerate the notebook from source:

```bash
cd /scratch/projects/CFP04/CFP04-CF-050/ces/super_right/github
python generate_notebooks.py
```

## Acknowledgement

The authors also gratefully acknowledge A/Prof. Sachin V. Jangam for providing the HDA Aspen HYSYS case study through the course `CN4205R`, and for his teaching in `CN4205 Pinch Analysis and Process Integration`, from which this work benefited.

## Contact

Questions, feedback, or collaboration ideas are welcome. Please contact: `ziyun_zhang@u.nus.edu`
