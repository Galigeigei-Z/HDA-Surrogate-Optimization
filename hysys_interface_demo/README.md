# HYSYS Interface Demo

This folder turns the CES Aspen HYSYS automation work into a teaching-oriented
demo for GitHub readers.

Contents:

- `notebooks/01_hysys_interface_tutorial.ipynb`: the main walkthrough notebook
- `hysys_demo/`: helper package used by the notebook

The notebook is designed to work in two modes:

- `mock` mode for GitHub readers and local testing without Aspen HYSYS
- `active_hysys` mode for a Windows machine with Aspen HYSYS open and
  `pywin32` installed

Topics covered in the notebook:

- what Python can read and write through the HYSYS COM interface
- how to connect to the active case
- how to discover real stream, energy-stream, and operation names safely
- how to map notebook parameters to stream and operation names
- how to do a minimal read-only sanity check before the first real write
- how to inspect streams before making changes
- how to apply a structured sample of changes to the flowsheet
- how to summarize energy-stream usage after a run
- how to retarget the same pattern to another HYSYS case
- how to debug the most common COM and naming failures
