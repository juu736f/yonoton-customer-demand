# Customer Demand Visualizer

This Python script generates hourly customer demand graphs from a `payments_captured.xlsx` file exported from the Yonoton Management Console.

## Features

- Cleans and parses timestamp fields.
- Aggregates payment amounts by hour.
- Produces bar chart images per date.
- Skips already processed dates.

## Requirements

- Python 3.8+
- Dependencies:
  - `pandas`
  - `matplotlib`
  - `openpyxl`
  - `tkinter` (standard library)

## Installation

```
pip install pandas matplotlib openpyxl
```

## Usage

1. Export the `payments_captured.xlsx` file from Yonoton.
2. Run the script:

```
python customer_demand_visualizer.py
```

3. Select the Excel file in the pop-up dialog.
4. Graphs are saved in `customer-demand-graphs/` as PNGs:

```
customer-demand-graphs/
├── 7.7.2025_demand.png
├── 8.7.2025_demand.png
```

## Output

* One image per date.
* Filename format: `D.M.YYYY_demand.png`.

## License

This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or distribute this software, either in source code form or as a compiled binary, for any purpose, commercial or non-commercial, and by any means.

For more information, please refer to [http://unlicense.org/](http://unlicense.org/)


