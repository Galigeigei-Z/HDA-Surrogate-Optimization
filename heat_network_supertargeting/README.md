# Heat Network Supertargeting

This folder preserves the original heat-integration part of the HDA workflow,
but restructures it into a standalone GitHub-friendly demo.

Contents:

- `supertargeting_demo.ipynb`: notebook walkthrough for workbook-driven pinch and
  supertargeting analysis
- `data/Input Sheet of 20.xlsx`: stream table used by the notebook
- `heat_network_demo/`: local helper package used by the notebook
- `images/`: pre-rendered figures from the original workflow

What this demo covers:

- reading thermal streams from the Excel workbook
- converting the workbook rows into `ThermalStream` objects
- running automatic supertargeting at a chosen `Delta Tmin`
- reviewing pinch temperatures, utility targets, and area intervals
- plotting composite-curve style summaries from the computed results

The notebook is now self-contained at the repository level. It no longer
depends on private path assumptions inside the original CES workspace.
