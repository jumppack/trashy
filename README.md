# Trashy Tracker

**Trash Rotation Tracker for KM, NP, PS, LS**

This project generates a printable trash rotation tracker spreadsheet. It tracks whose turn it is to take out the trash among the four participants: KM, NP, PS, and LS.

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
└── output/                # Generated files
    └── Trash_Rotation_Printable.xlsx
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

The output file `Trash_Rotation_Printable.xlsx` will be generated in the `output/` directory.

## Git Repository Setup

To set up this project on GitHub:

1.  **Create a New Repository on GitHub:**
    *   Go to [https://github.com/new](https://github.com/new).
    *   Name the repository `trashy-tracker` (or your preferred name).
    *   Do **not** check "Initialize this repository with a README", .gitignore, or license (we already have them).
    *   Click "Create repository".

2.  **Push to GitHub:**
    Run the following commands in your terminal (inside the project folder):

    ```bash
    git add .
    git commit -m "Initial commit: Project organization and tracker script"
    git branch -M main
    git remote add origin https://github.com/YOUR_USERNAME/trashy-tracker.git
    git push -u origin main
    ```
    *(Replace `YOUR_USERNAME` with your actual GitHub username)*.

## How it Works

The script uses `pandas` and `xlsxwriter` to create a formatted Excel sheet. It creates a simple rotation flow:
`KM -> NP -> PS -> LS`

Find the first empty box and mark it done!

## Methodology

This tracker is based on the **Harmonious Household** strategy, aiming to reduce mental load and decision fatigue through visual accountability. See [docs/resources/harmonious_household.md](docs/resources/harmonious_household.md) for more details.
