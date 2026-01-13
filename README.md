# Trashy Tracker

**Trash Rotation Tracker for KM, NP, PS, LS (v2)**

This project generates a high-density, printable trash rotation tracker designated for A4 printing. It utilizes the **"Dent Method"** for marking completion without pens.

## Project Structure

```
Trashy Tracker/
├── .gitignore             # Git ignore file
├── README.md              # Project documentation
├── requirements.txt       # Python dependencies
├── src/
│   └── generate_tracker.py # Script to generate the tracker
├── notebooks/             # Archived notebooks
│   ├── trash_tracker_v1_using_matplotlib.ipynb
│   └── trash_tracker_v2_using_pandas.ipynb
├── docs/                  # Documentation and resources
│   └── resources/
│       └── harmonious_household.md # Strategy details
└── out_v2/                # Generated files (v2)
    └── Trash_Rotation_Printable_v2.xlsx
```

## Setup and Installation

1.  **Install Dependencies:**
    It is recommended to use a virtual environment.
    ```bash
    pip install -r requirements.txt
    ```

## Usage

Run the generation script to create the Excel file:

```bash
python src/generate_tracker.py
```

The output file `Trash_Rotation_Printable_v2.xlsx` will be generated in the `out_v2/` directory. Open this file and **Print to A4** (Scaling is set to Fit on One Page).

## Git Repository Setup

The project is hosted at: [https://github.com/jumppack/trashy](https://github.com/jumppack/trashy)

To push updates:
```bash
git add .
git commit -m "Your commit message"
git push
```

## How it Works

`KM -> NP -> PS -> LS -> (Next Row)`

**The Dent Method:**
Instead of looking for a pen, simply find the first empty bracket box `[     ]` and press your house key firmly into the paper to create a visible dent/shadow.

## Methodology

This tracker is based on the **Harmonious Household** strategy, aiming to reduce mental load and decision fatigue through visual accountability. See [docs/resources/harmonious_household.md](docs/resources/harmonious_household.md) for more details.
