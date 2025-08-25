# Fusion 360 2D Export Add-in

A specialized Fusion 360 add-in for bulk exporting 2D data (drawings) from all your designs with configurable file types and export formats.

## Features

### Configurable File Types
- **Drawings**: Export technical drawings
 

### Folder Selection
- **Export All**: Process all folders in all projects (default behavior)
- **Selective Export**: Choose specific folders to export from
- **Multi-Selection**: Select multiple folders across different projects

### Supported Export Formats
- **DXF**: Drawing Exchange Format
- **PDF**: Portable Document Format


## Installation (from GitHub)

There are two easy ways to install this add-in from GitHub.

### Option A — Download ZIP (no Git required)
1. Download the repository as a ZIP from the GitHub page (Code → Download ZIP).
2. Extract the ZIP.
3. Create a new folder for the add-in (e.g., `Fusion_2D_Export`).
4. Copy these three files into that folder:
   - `Fusion 360 2D Export.py`
   - `Fusion 360 2D Export.manifest`
   - `folder_browser.html`
5. Move that folder into your Fusion scripts/add-ins directory:
   - Windows (newer Fusion): `C:\Users\<you>\AppData\Roaming\Autodesk\Autodesk Fusion\API\Scripts\`
   - Windows (legacy): `C:\Users\<you>\AppData\Roaming\Autodesk\Autodesk Fusion 360\API\Scripts\`
   - macOS: `~/Library/Application Support/Autodesk/Autodesk Fusion/API/Scripts/`
6. Start Fusion → go to **Utilities** (or **Tools**) → **Add-Ins** (or **Scripts/Add-Ins**).
7. In the Scripts tab, click the **+** or **Browse** button, select the folder you created, and add it.

### Option B — Git clone (keeps it up to date)
```bash
git clone https://github.com/<your-username>/<your-repo>.git
```
Then copy the same three files into a folder under your Fusion Scripts directory as described above, or clone directly into a folder under `.../API/Scripts/`.

### Updating
- If you installed via ZIP: repeat Option A and replace the three files in your folder.
- If you installed via Git: `git pull` in your clone and restart Fusion.

## Usage

1. Run the "Fusion 360 2D Export" script from the Scripts/Add-ins panel
2. Configure your export settings in the dialog:
   - **File Types Tab**: Select what to export (drawings)
   - **Export Formats Tab**: Choose output formats (DXF, PDF)
   - **Folder Selection Tab**: Choose to export all folders or select specific ones
   - **Performance Tab**: Enable cache clearing for better memory management
3. Click **OK** to proceed
4. Select the output directory where exported files will be saved
5. The tool will process your selected folders and export the 2D data

## Output Structure

The exported files are organized in a hierarchical structure:
```
Output Folder/
├── Hub [Hub Name]/
│   ├── Project [Project Name]/
│   │   ├── [Design Name]/
│   │   │   └── Drawings/
│   │   │       ├── Drawing1.pdf
│   │   │       └── Drawing1.dxf
│   │   └── [Another Design]/
│   └── [Another Project]/
└── export_2d.log
```



## Troubleshooting

- Check the `export_2d.log` file in your output directory for detailed error information

