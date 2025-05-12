# Cluster and Form

A powerful tool for organizing university activities by automatically forming balanced teams while respecting friend preferences and team constraints.

## 🎯 Problem Statement

Organizing university activities often requires forming teams with multiple constraints:
- Teams should have roughly equal sizes
- Equal distribution of genders for fair play
- Mix of students from different university years
- Respect friend preferences when possible

Manually forming teams while considering all these constraints is time-consuming and error-prone. This tool automates the entire process, saving hours of manual work.

## ✨ Features

- **Excel Data Import**: Load participant data including names, phone numbers, types (leader/subleader), gender, year (dashID), and friend preferences
- **Smart Name Matching**: Automatically matches friend recommendations even with partial or slightly incorrect names
- **Interactive Friend Management**: Edit and verify friend connections through an intuitive table interface
- **Visual Cluster Analysis**: Interactive network graph visualization of friend clusters using pyvis and vis.js
- **Dynamic Graph Editing**: Add, delete, or modify connections directly in the graph view
- **Intelligent Team Formation**: Uses evolutionary programming to create balanced teams while respecting all constraints
- **Excel Export**: Generates color-coded Excel files with team assignments and cluster information

## 🛠️ Technical Stack

### Core Libraries
- **PyQt5**: GUI framework
  - QtWebEngineWidgets: For web-based graph visualization
  - QtWebChannel: For JavaScript-Python communication
- **pandas**: Data manipulation and Excel file handling
- **openpyxl**: Excel file generation and styling
- **pyvis**: Network graph visualization
- **numpy**: Numerical computations
- **lxml**: XML processing for graph data

### Key Functions

#### Data Management
- `load_excel_file()`: Imports participant data from Excel
- `get_matches()`: Smart name matching algorithm
- `edit_friend_column()`: Interactive friend connection management

#### Clustering
- `assign_clusters()`: DFS-based cluster formation
- `split_large_clusters()`: Ensures manageable cluster sizes

#### Team Formation
- `distribute()`: Initial team distribution
- `feasibility_check()`: Validates team constraints
- `fitness()`: Calculates team distribution score
- `generate_solutions()`: Evolutionary algorithm for optimal team formation

#### Visualization
- `create_graph()`: Generates interactive network visualization
- `export_teams_to_excel()`: Creates color-coded team assignments

## 📸 Screenshots

<<<<<<< HEAD
The following screenshots are required for the application to function properly. Place them in the `images` directory:

1. `images/sampleInput2.png`: Example of required Excel input format
   - Shows the expected layout of the input Excel file
   - Required for the initial file loading dialog

2. `images/cluster.png`: Application icon
   - Used as the window icon
   - Required for proper application branding

3. `images/main_interface.png`: Main application interface
   - Shows the friend matching table
   - Demonstrates the interactive friend management system

4. `images/graph_view.png`: Network graph visualization
   - Shows the interactive cluster visualization
   - Demonstrates the graph editing capabilities

5. `images/teams_output.png`: Example of generated team Excel file
   - Shows the color-coded team assignments
   - Demonstrates the cluster coloring system

### Image Requirements
- All images should be placed in the `images` directory at the root of the project
- The application icon (`cluster.png`) should be in PNG format
- `sampleInput2.png` should be clear enough to show the Excel column structure
- Screenshots should be taken at a reasonable resolution (recommended: 1920x1080 or higher)
=======
[Add screenshots in the following order:]

1. `images/sampleInput2.png`: Example of required Excel input format
2. `images/cluster.png`: Application icon
3. `images/main_interface.png`: Main application interface showing the friend matching table
4. `images/graph_view.png`: Network graph visualization of clusters
5. `images/teams_output.png`: Example of generated team Excel file
>>>>>>> 54517c622f25632d9967d554b02f67389415b848

## 🚀 Installation

1. Clone the repository
2. Install required packages:
```bash
pip install -r requirements.txt
```

## 💻 Usage

1. Prepare your Excel file with the following columns:
   - Name
   - Dash ID
   - Gender
   - Type
   - Number
   - Friend 1
   - Friend 2

2. Run the application:
```bash
python cluster-and-form.py
```

3. Load your Excel file and verify friend matches
4. Use the graph view to visualize and edit clusters
5. Generate teams and export to Excel

## 📦 Building Executable

<<<<<<< HEAD
To create a standalone executable, ensure all required images are in place and run:
```bash
pyinstaller --onefile --windowed ^
--hidden-import tkinter ^
--hidden-import openpyxl ^
--hidden-import pyvis ^
--hidden-import pandas ^
--hidden-import numpy ^
--hidden-import lxml ^
--hidden-import collections ^
--hidden-import copy ^
--hidden-import random ^
--hidden-import math ^
--hidden-import sys ^
--hidden-import PyQt5 ^
--hidden-import PyQt5.QtCore ^
--hidden-import PyQt5.QtGui ^
--hidden-import PyQt5.QtWidgets ^
--hidden-import PyQt5.QtWebEngineWidgets ^
--hidden-import PyQt5.QtWebChannel ^
--add-data "images/PYF-Logo.jpg;." ^
--add-data "path/to/pyvis/templates;pyvis/templates" ^
--icon "images/cluster.png" ^
cluster-and-form.py
```

Note: Update the paths in the PyInstaller command according to your system's configuration.

=======
To create a standalone executable:
```bash
pyinstaller --onefile --windowed --hidden-import tkinter --hidden-import openpyxl --hidden-import pyvis --hidden-import pandas --hidden-import numpy --hidden-import lxml --hidden-import collections --hidden-import copy --hidden-import random --hidden-import math --hidden-import sys --hidden-import PyQt5 --hidden-import PyQt5.QtCore --hidden-import PyQt5.QtGui --hidden-import PyQt5.QtWidgets --hidden-import PyQt5.QtWebEngineWidgets --hidden-import PyQt5.QtWebChannel --add-data "path/to/PYF-Logo.jpg;." --add-data "path/to/pyvis/templates;pyvis/templates" --icon "path/to/PYF-Logo.jpg" cluster-and-form.py
```

>>>>>>> 54517c622f25632d9967d554b02f67389415b848
## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📝 License

[Add your chosen license here]

## 👤 Author

<<<<<<< HEAD
[Your name/contact information] 
=======
[Your name/contact information]
>>>>>>> 54517c622f25632d9967d554b02f67389415b848
