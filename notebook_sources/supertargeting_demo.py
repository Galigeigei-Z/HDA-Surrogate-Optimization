from __future__ import annotations


def _markdown_cell(text: str) -> dict[str, object]:
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": text.splitlines(keepends=True),
    }


def _code_cell(code: str) -> dict[str, object]:
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": code.splitlines(keepends=True),
    }


def build_notebook() -> dict[str, object]:
    cells = [
        _markdown_cell(
            "# Supertargeting Demo\n\n"
            "This notebook demonstrates a minimal workflow for reading heat-integration stream data from Excel and "
            "running automatic supertargeting calculations."
        ),
        _markdown_cell(
            "## What This Demo Covers\n\n"
            "1. Read the local Excel stream sheet\n"
            "2. Convert rows into `ThermalStream` objects\n"
            "3. Run supertargeting for a selected $\\Delta T_{\\min}$\n"
            "4. Inspect the stream table, summary metrics, area intervals, composite curve, and balanced composite curve"
        ),
        _code_cell(
            "from pathlib import Path\n"
            "import sys\n\n"
            "import matplotlib.pyplot as plt\n"
            "import pandas as pd\n\n"
            "from IPython.display import Image, display\n\n"
            "PROCESS_ONLY_UTILITY_NAMES = {'Hot Oil', 'HP Steam', 'Cooling Water', 'Refrigerant 2', 'Fired Heat (1000)'}\n\n"
            "NOTEBOOK_DIR = Path.cwd().resolve()\n"
            "SUPER_RIGHT_DIR = NOTEBOOK_DIR.parent\n"
            "ROOT = SUPER_RIGHT_DIR.parent\n"
            "SUPERTARGETING_ROOT = ROOT / '06_HYSYS_Supertargeting'\n\n"
            "if str(ROOT) not in sys.path:\n"
            "    sys.path.insert(0, str(ROOT))\n"
            "if str(SUPERTARGETING_ROOT) not in sys.path:\n"
            "    sys.path.insert(0, str(SUPERTARGETING_ROOT))\n\n"
            "from tex_tables.generate_supertargeting_area_target import parse_xlsx_rows, stream_rows_from_workbook\n"
            "from hysys_utils.supertargeting import _rows_to_streams, curve_plot_records, run_supertargeting\n"
        ),
        _code_cell(
            "INPUT_XLSX = NOTEBOOK_DIR / 'Input Sheet of 20.xlsx'\n"
            "IMAGE_DIR = NOTEBOOK_DIR / 'images'\n"
            "DELTA_TMIN_C = 20.0\n\n"
            "INPUT_XLSX"
        ),
        _code_cell(
            "workbook_rows = parse_xlsx_rows(INPUT_XLSX)\n"
            "stream_rows = stream_rows_from_workbook(workbook_rows)\n\n"
            "streams_df = pd.DataFrame(\n"
            "    stream_rows,\n"
            "    columns=['Stream Information', 'Supply Temperature (C)', 'Target Temperature (C)', 'Heat Load (kW)', 'U (kW/m2-K)', 'CP (kW/K)'],\n"
            ")\n"
            "streams_df.index = streams_df.index + 1\n"
            "streams_df"
        ),
        _code_cell(
            "streams = _rows_to_streams(workbook_rows)\n"
            "result = run_supertargeting(streams, delta_tmin_c=DELTA_TMIN_C)\n\n"
            "summary_df = pd.DataFrame(\n"
            "    [\n"
            "        {'Metric': 'Delta Tmin (C)', 'Value': DELTA_TMIN_C},\n"
            "        {'Metric': 'Pinch hot temperature (C)', 'Value': result.pinch_hot_c},\n"
            "        {'Metric': 'Pinch cold temperature (C)', 'Value': result.pinch_cold_c},\n"
            "        {'Metric': 'Minimum hot utility (kW)', 'Value': result.minimum_hot_utility_kw},\n"
            "        {'Metric': 'Minimum cold utility (kW)', 'Value': result.minimum_cold_utility_kw},\n"
            "        {'Metric': 'Minimum area (m2)', 'Value': result.minimum_area_m2},\n"
            "    ]\n"
            ")\n"
            "summary_df"
        ),
        _markdown_cell(
            "## Total Results Table\n\n"
            "This consolidated table summarizes the current `\\Delta T_{\\min}=20^\\circ C` case for both all-stream and process-only views."
        ),
        _code_cell(
            "process_streams = [stream for stream in streams if stream.name not in PROCESS_ONLY_UTILITY_NAMES]\n"
            "assert not any(stream.name in PROCESS_ONLY_UTILITY_NAMES for stream in process_streams)\n"
            "process_result = run_supertargeting(process_streams, delta_tmin_c=DELTA_TMIN_C)\n\n"
            "total_results_df = pd.DataFrame(\n"
            "    [\n"
            "        {'Basis': 'All streams', 'Metric': 'Stream count', 'Value': len(streams)},\n"
            "        {'Basis': 'All streams', 'Metric': 'Pinch hot temperature (C)', 'Value': result.pinch_hot_c},\n"
            "        {'Basis': 'All streams', 'Metric': 'Pinch cold temperature (C)', 'Value': result.pinch_cold_c},\n"
            "        {'Basis': 'All streams', 'Metric': 'Minimum hot utility (kW)', 'Value': result.minimum_hot_utility_kw},\n"
            "        {'Basis': 'All streams', 'Metric': 'Minimum cold utility (kW)', 'Value': result.minimum_cold_utility_kw},\n"
            "        {'Basis': 'All streams', 'Metric': 'Minimum area (m2)', 'Value': result.minimum_area_m2},\n"
            "        {'Basis': 'Process only', 'Metric': 'Stream count', 'Value': len(process_streams)},\n"
            "        {'Basis': 'Process only', 'Metric': 'Pinch hot temperature (C)', 'Value': process_result.pinch_hot_c},\n"
            "        {'Basis': 'Process only', 'Metric': 'Pinch cold temperature (C)', 'Value': process_result.pinch_cold_c},\n"
            "        {'Basis': 'Process only', 'Metric': 'Minimum hot utility (kW)', 'Value': process_result.minimum_hot_utility_kw},\n"
            "        {'Basis': 'Process only', 'Metric': 'Minimum cold utility (kW)', 'Value': process_result.minimum_cold_utility_kw},\n"
            "        {'Basis': 'Process only', 'Metric': 'Minimum area (m2)', 'Value': process_result.minimum_area_m2},\n"
            "    ]\n"
            ")\n"
            "total_results_df"
        ),
        _code_cell(
            "intervals_df = pd.DataFrame(\n"
            "    [\n"
            "        {\n"
            "            'Delta H Interval (kW)': f'{interval.enthalpy_start_kw:.1f}--{interval.enthalpy_end_kw:.1f}',\n"
            "            'Th_out (C)': interval.hot_out_c,\n"
            "            'Th_in (C)': interval.hot_in_c,\n"
            "            'Tc_in (C)': interval.cold_in_c,\n"
            "            'Tc_out (C)': interval.cold_out_c,\n"
            "            'Hot Streams': list(interval.hot_stream_ids),\n"
            "            'Cold Streams': list(interval.cold_stream_ids),\n"
            "            'LMTD (C)': interval.log_mean_temperature_difference_c,\n"
            "            'Area (m2)': interval.area_m2,\n"
            "        }\n"
            "        for interval in result.area_intervals\n"
            "    ]\n"
            ")\n"
            "intervals_df"
        ),
        _code_cell(
            "records_bcc = pd.DataFrame(curve_plot_records(streams, result))\n"
            "records_cc = pd.DataFrame(curve_plot_records(process_streams, process_result))\n\n"
            "pd.concat(\n"
            "    [\n"
            "        records_cc.assign(source='process_only_cc'),\n"
            "        records_bcc.assign(source='all_streams_bcc'),\n"
            "    ],\n"
            "    ignore_index=True,\n"
            ").head()"
        ),
        _code_cell(
            "fig, axes = plt.subplots(1, 2, figsize=(12, 5), constrained_layout=True)\n\n"
            "plot_specs = [\n"
            "    (records_cc, 'composite_curve', 'Composite Curve (process streams only)'),\n"
            "    (records_bcc, 'bcc_curve', 'Balanced Composite Curve (including utility streams)'),\n"
            "]\n\n"
            "for ax, (records_plot, plot_name, title) in zip(axes, plot_specs):\n"
            "    hot_rows = records_plot[(records_plot['plot_name'] == plot_name) & (records_plot['curve_name'] == 'hot_curve')]\n"
            "    cold_rows = records_plot[(records_plot['plot_name'] == plot_name) & (records_plot['curve_name'] == 'cold_curve')]\n"
            "    ax.plot(hot_rows['enthalpy_kw'], hot_rows['temperature_c'], color='#c62828', linewidth=2.2, label='Hot Composite Curve')\n"
            "    ax.plot(cold_rows['enthalpy_kw'], cold_rows['temperature_c'], color='#1565c0', linewidth=2.2, label='Cold Composite Curve')\n"
            "    ax.set_title(title)\n"
            "    ax.set_xlabel('Enthalpy (kW)')\n"
            "    ax.set_ylabel('Temperature (C)')\n"
            "    ax.grid(True, color='#d9d9d9', linewidth=0.8)\n"
            "    ax.legend(frameon=True)\n\n"
            "plt.show()\n"
        ),
        _markdown_cell(
            "## Rendered Figures\n\n"
            "The notebook also shows the pre-rendered figures copied into this `github` demo folder."
        ),
        _code_cell(
            "figure_map = {\n"
            "    'Balanced composite curve': IMAGE_DIR / 'bcc_all_streams_dtmin20.png',\n"
            "    'Partitioned BCC': IMAGE_DIR / 'bcc_all_streams_partitioned_dtmin20.png',\n"
            "    'Composite curve without utility streams': IMAGE_DIR / 'composite_curve_process_only_dtmin20.png',\n"
            "    'ACC / TOC / TAC vs Delta Tmin': IMAGE_DIR / 'dtmin_economics_aligned_bcc.png',\n"
            "}\n\n"
            "pd.DataFrame({'Figure': list(figure_map.keys()), 'Path': [str(path) for path in figure_map.values()]})"
        ),
        _code_cell(
            "for title, path in figure_map.items():\n"
            "    print(title)\n"
            "    display(Image(filename=str(path)))\n"
        ),
        _markdown_cell(
            "## Notes\n\n"
            "- The demo uses the local copy of `Input Sheet of 20.xlsx` in this folder.\n"
            "- Change `DELTA_TMIN_C` and rerun the notebook cells to test another case.\n"
            "- The same backend can be reused later for publishing a cleaner standalone GitHub example.\n"
            "- The `images/` subfolder is intended to keep the demo notebook self-contained."
        ),
        _markdown_cell(
            "## Acknowledgement\n\n"
            "The authors also gratefully acknowledge A/Prof. Sachin V. Jangam for providing the HDA Aspen HYSYS case study through the course `CN4205R`, and for his teaching in `CN4205 Pinch Analysis and Process Integration`, from which this work benefited."
        ),
        _markdown_cell(
            "## Contact\n\n"
            "Questions, feedback, or collaboration ideas are welcome. Please contact: `ziyun_zhang@u.nus.edu`"
        ),
    ]

    return {
        "cells": cells,
        "metadata": {
            "kernelspec": {
                "display_name": "Python 3",
                "language": "python",
                "name": "python3",
            },
            "language_info": {
                "name": "python",
                "version": "3.11",
            },
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }
