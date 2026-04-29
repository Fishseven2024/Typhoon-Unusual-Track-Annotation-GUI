# Typhoon-Unusual-Track-Annotation-GUI
A research-oriented toolkit for visual annotation and validation of unusual tropical cyclone tracks.

# Typhoon Unusual Track Statistic GUI

A lightweight GUI tool for manually marking, reviewing, and statistically summarizing unusual tropical cyclone tracks.

This project was developed for research on unusual typhoon tracks in the Western North Pacific. The tool supports single-track visualization, interactive segment marking, automatic statistics, progress persistence, and reproducible export of marked results.
![Uploading Quicker_20260429_232928.png…]()

## Features

- Visualize tropical cyclone tracks on a map
- Manually mark unusual track segments
- Support point-based and segment-based interactive selection
- Filter tracks by year, category, prediction labels, or custom fields
- Save marking progress with SQLite
- Export marked tracks and statistics
- Generate high-resolution PNG figures
- Export segment-level CSV files
- Export track-level TXT and XLSX files
- Generate summary XLSX tables for completed marked tracks

## Research Use Case

This GUI is designed for human-assisted construction and verification of an unusual typhoon track dataset.

The main workflow is:

1. Load tropical cyclone best-track data
2. Inspect each track visually
3. Mark local unusual segments
4. Save marking progress
5. Export figures and statistical summaries
6. Use the marked results for dataset construction and model validation

## Data Availability

The research dataset used in this project is not included in this repository.

The current version of the software can be adapted to public best-track datasets such as IBTrACS. The manually curated typhoon track dataset associated with the author's research will be released after publication when appropriate.

## Repository Contents

```text
.
├── Statistic_GUI_2.1.py      # Main GUI program
├── README.md                 # Project description
├── LICENSE                   # Open-source license
└── requirements.txt          # Python dependencies, if provided
````

## Input Data

The GUI currently expects two main input files:

1. A tropical cyclone track CSV file
2. A track attribute XLSX file

The expected track CSV should contain at least the following fields:

```text
SID
SEASON
NAME
BASIN
LAT
LON
USA_WIND
ISO_TIME
```

The XLSX attribute table should contain at least:

```text
SID
```

Additional columns can be used for filtering and classification support.

## Output Files

The program can generate several types of output:

```text
Statistic_GUI_Output/
├── db/
│   ├── statistic_gui_progress.db
│   └── operation_history.txt
├── marked_plots/
├── segment_csv/
├── track_txt/
├── track_xlsx/
└── summary/
```

Main output types include:

* Marked track PNG figures
* Segment-level CSV files
* Track-level TXT summaries
* Track-level XLSX files
* Completed-track statistical summaries
* SQLite progress database
* Operation history logs

## Installation

Create a Python environment and install the required packages.

Typical dependencies include:

```text
numpy
pandas
matplotlib
cartopy
openpyxl
```

The GUI also uses the Python standard library modules `tkinter`, `sqlite3`, `pathlib`, and others.

## Usage

1. Clone this repository:

```bash
git clone https://github.com/your-username/your-repository-name.git
cd your-repository-name
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Modify the fixed path configuration in the main script:

```python
WORKSPACE_ROOT = Path(...)
TRACK_CSV_PATH = Path(...)
ATTRIBUTE_XLSX_PATH = Path(...)
OUTPUT_ROOT = ...
```

4. Run the GUI:

```bash
python Statistic_GUI_2.1.py
```

## Notes

This project focuses on manual visual marking and statistical organization of unusual typhoon track segments. It does not include a target detection or automatic recognition model inside the GUI.

The GUI is intended as a research-support tool for dataset construction, human review, and reproducible result export.

## License

This software is released under the MIT License.

The license only applies to the source code in this repository. Research data, manually labeled datasets, figures, and unpublished experimental results are not included unless explicitly stated.

## Citation

If this tool or its derived dataset is useful for your research, please cite the related paper after it becomes available.

## Author

Developed for research on unusual tropical cyclone tracks and typhoon path dataset construction.
## Attention！！！
This repository releases the software tools under the MIT License. 
The typhoon track dataset used in the associated research is not included in this repository and will be released after publication when appropriate.
