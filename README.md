# Fusion 360 2D Export Add-in

A specialized Fusion 360 add-in for bulk exporting 2D data (sketches and drawings) from all your designs with configurable file types and export formats.

## Features

### Configurable File Types
- **Sketches**: Export all sketches from designs
- **Drawings**: Export technical drawings
- **Construction Planes**: Export construction geometry (planned)
- **Work Features**: Export work points, axes, and planes (planned)

### Folder Selection
- **Export All**: Process all folders in all projects (default behavior)
- **Selective Export**: Choose specific folders to export from
- **Multi-Selection**: Select multiple folders across different projects

### Supported Export Formats
- **DXF**: AutoCAD Drawing Exchange Format
- **DWG**: AutoCAD Drawing Format (exports as DXF)
- **PDF**: Portable Document Format (drawings only)
- **SVG**: Scalable Vector Graphics (planned)
- **PNG**: Portable Network Graphics (planned)
- **JPEG**: Joint Photographic Experts Group (planned)

## Installation

1. Download or clone this repository
2. Copy the entire `bulkExport` folder to your Fusion 360 scripts directory
3. Open Fusion 360
4. Go to **Tools** → **Scripts/Add-ins**
5. Click the **+** button and select the `bulkExport` folder
6. The "Fusion 360 2D Export" script will appear in the list

## Usage

1. Run the "Fusion 360 2D Export" script from the Scripts/Add-ins panel
2. Configure your export settings in the dialog:
   - **File Types Tab**: Select what to export (sketches, drawings, etc.)
   - **Export Formats Tab**: Choose output formats (DXF, PDF, etc.)
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
│   │   │   ├── Sketches/
│   │   │   │   ├── Sketch1.dxf
│   │   │   │   └── Sketch2.dxf
│   │   │   └── Drawings/
│   │   │       ├── Drawing1.pdf
│   │   │       └── Drawing1.dxf
│   │   └── [Another Design]/
│   └── [Another Project]/
└── export_2d.log
```

## Key Improvements Over Original

### Enhanced Configuration
- **Interactive Dialog**: Choose exactly what to export and in which formats
- **File Type Selection**: Focus on specific 2D elements you need
- **Format Options**: Multiple export formats for different use cases

### 2D-Focused Design
- **Optimized Performance**: Only processes 2D data, faster execution
- **Specialized Handling**: Better support for sketches and drawings
- **Organized Output**: Clear separation of sketches and drawings

### Better User Experience
- **Progress Tracking**: Visual progress dialog with detailed status
- **Comprehensive Logging**: Detailed log file for troubleshooting
- **Error Handling**: Graceful handling of export failures

## Limitations

- Some export formats are not yet fully implemented (SVG, PNG, JPEG)
- DWG format exports as DXF due to API limitations
- Requires Fusion 360 with appropriate licensing for drawing features

## Troubleshooting

- Check the `export_2d.log` file in your output directory for detailed error information
- Ensure you have the necessary permissions to write to the selected output directory
- Some designs may not have sketches or drawings to export

## Based On

This add-in is modified from the excellent [Fusion 360 Total Exporter](https://github.com/Jnesselr/fusion-360-total-exporter) by Justin Nesselrotte, specialized for 2D data export with enhanced configuration options.
