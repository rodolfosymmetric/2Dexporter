from __future__ import with_statement

import adsk.core, adsk.fusion, adsk.cam, traceback
import os
import re
import shutil
import tempfile
import logging
import json
import threading
import time


class Logger(object):
    """Simple logger for Fusion 360 add-ins"""
    def __init__(self, name):
        self.name = name
        
    def info(self, message):
        print("[INFO] {}: {}".format(self.name, message))
        
    def warning(self, message):
        print("[WARNING] {}: {}".format(self.name, message))
        
    def error(self, message):
        print("[ERROR] {}: {}".format(self.name, message))


class TwoDExport(object):
    def __init__(self, app):
        self.app = app
        self.ui = self.app.userInterface
        self.data = self.app.data
        self.documents = self.app.documents
        self.log = Logger("Fusion 360 2D Export")
        self.num_issues = 0
        self.was_cancelled = False
        
        # Configuration options
        self.file_types_to_export = []
        self.export_formats = []
        self.selected_folders = []  # List of selected folder paths to export
        
        # Available file types to export
        self.available_file_types = {
            'sketches': 'Sketches',
            'drawings': 'Drawings', 
            'construction_planes': 'Construction Planes',
            'work_features': 'Work Features'
        }
        
        # Available export formats (primary formats first)
        self.available_formats = {
            'dxf': 'DXF (AutoCAD Drawing Exchange)',
            'dwg': 'DWG (AutoCAD Drawing)', 
            'pdf': 'PDF (Portable Document Format)',
            'svg': 'SVG (Scalable Vector Graphics) - Planned',
            'png': 'PNG (Portable Network Graphics) - Planned',
            'jpg': 'JPEG (Joint Photographic Experts Group) - Planned'
        }
        
        # Primary supported formats
        self.primary_formats = ['dxf', 'dwg', 'pdf']
        
        # Cache management option
        self.clear_cache_after_each_file = False
        
        # Folder selection option
        self.export_all_folders = True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def run(self, context):
        # Show configuration dialog first
        if not self._show_configuration_dialog():
            return
            
        self.ui.messageBox(
            "Exporting 2D data will take a while, depending on how many files you have.\n\n"
            "You won't be able to do anything else during the export process.\n\n"
            "The tool will export: {}\n"
            "In formats: {}\n\n"
            "NOTE: The export will include ALL files in the selected folder AND all its sub-folders.".format(
                ", ".join([self.available_file_types[ft] for ft in self.file_types_to_export]),
                ", ".join([self.available_formats[fmt] for fmt in self.export_formats])
            )
        )

        output_path = self._ask_for_output_path()

        if output_path is None:
            return

        file_handler = FileHandler(os.path.join(output_path, 'export_2d.log'))
        file_handler.setFormatter(Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.log.addHandler(file_handler)

        self.log.info("Starting 2D export!")
        self.log.info("File types to export: {}".format(self.file_types_to_export))
        self.log.info("Export formats: {}".format(self.export_formats))

        self._export_2d_data(output_path)

        self.log.info("Done exporting 2D data!")

        if self.was_cancelled:
            self.ui.messageBox("Export cancelled!")
        elif self.num_issues > 0:
            self.ui.messageBox("The exporting process ran into {} issue{}. Please check the log for more information".format(
                self.num_issues,
                "s" if self.num_issues > 1 else ""
            ))
        else:
            self.ui.messageBox("2D Export finished successfully!")

    def _show_configuration_dialog(self):
        """Show 3-step configuration for 2D drawings export"""
        try:
            # Step 1: Select source folder from Fusion 360 cloud
            self.ui.messageBox(
                "Step 1 of 3: Select Cloud Folder\n\n"
                "Next, you'll browse your Fusion 360 cloud storage to select the folder containing drawings to export.",
                "2D Drawings Export"
            )
            
            cloud_folder = self._ask_for_cloud_folder()
            if cloud_folder is None:
                return False
            
            self.selected_cloud_folder = cloud_folder
            
            # Step 2: Export format selection
            format_result = self.ui.messageBox(
                "Step 2 of 3: Select Export Format\n\n"
                "Click 'Yes' for PDF\n"
                "Click 'No' for DXF\n"
                "Click 'Cancel' to exit",
                "Select Export Format", 
                adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
            )
            
            if format_result == adsk.core.DialogResults.DialogYes:
                self.export_formats = ['pdf']
            elif format_result == adsk.core.DialogResults.DialogNo:
                self.export_formats = ['dxf']
            elif format_result == adsk.core.DialogResults.DialogCancel:
                return False  # User cancelled
            else:
                return False
            
            # Ask if user wants both formats
            both_formats = self.ui.messageBox(
                "Do you also want to export in the other format?\n\n"
                "Click 'Yes' to export in both PDF and DXF\n"
                "Click 'No' to continue with {} only".format(self.export_formats[0].upper()),
                "Export Both Formats?",
                adsk.core.MessageBoxButtonTypes.YesNoButtonType
            )
            
            if both_formats == adsk.core.DialogResults.DialogYes:
                self.export_formats = ['pdf', 'dxf']
            
            # Step 3: Output folder selection (handled in main run method)
            self.ui.messageBox(
                "Step 3 of 3: Select Output Folder\n\n"
                "Next, you'll select where to save the exported drawings.",
                "Select Output Location"
            )
            
            # Set defaults for 2D drawings export
            self.file_types_to_export = ['drawings']  # Focus only on drawings (blueprints with dimensions)
            self.export_all_folders = False  # We're using selected folder
            self.clear_cache_after_each_file = True
            self.selected_folders = []
            
            return True
            
        except Exception as ex:
            self.ui.messageBox("Error in configuration: {}".format(str(ex)))
            return False

    def _ask_for_cloud_folder(self):
        """Ask user to select a folder from Fusion 360 cloud storage using HTML palette"""
        try:
            return self._show_html_folder_browser()
        except Exception as ex:
            self.ui.messageBox("Error browsing cloud folders: {}".format(str(ex)))
            return None

    def _show_html_folder_browser(self):
        """Show HTML palette for folder browsing"""
        try:
            import os
            
            # Get the path to the HTML file
            script_dir = os.path.dirname(os.path.realpath(__file__))
            html_file = os.path.join(script_dir, 'folder_browser.html')
            
            if not os.path.exists(html_file):
                self.ui.messageBox("HTML file not found: {}".format(html_file))
                return None
            
            # Convert to proper file URL format for Fusion 360
            html_file = html_file.replace('\\', '/')
            if not html_file.startswith('file:///'):
                html_file = 'file:///' + html_file
            
            # Remove existing palette if it exists
            existing_palette = self.ui.palettes.itemById('FolderBrowser')
            if existing_palette:
                existing_palette.deleteMe()
            
            # Create and show the palette
            palette = self.ui.palettes.add(
                'FolderBrowser',
                'Select Cloud Folder', 
                html_file,
                True,  # showCloseButton
                True,  # isResizable
                True,  # isVisible
                400,   # width
                600    # height
            )
            
            # Set up event handlers for communication with HTML
            self._setup_palette_handlers(palette)
            
            # Initialize the HTML with data after a short delay
            self._initialize_html_data(palette)
            
            # Wait for user selection
            self.selected_folder_result = None
            self.palette_closed = False
            
            # Keep the palette open until user makes a selection or cancels
            while not self.palette_closed and self.selected_folder_result is None:
                adsk.doEvents()
            
            # Clean up
            if palette:
                palette.deleteMe()
            
            return self.selected_folder_result
            
        except Exception as ex:
            self.ui.messageBox("Error showing HTML folder browser: {}".format(str(ex)))
            return None

    def _setup_palette_handlers(self, palette):
        """Set up event handlers for HTML palette communication"""
        try:
            # Add event handler for HTML requests
            onHTMLEvent = HTMLEventHandler(self)
            palette.incomingFromHTML.add(onHTMLEvent)
            
            # Add event handler for palette closing
            onClosed = PaletteClosedHandler(self)
            palette.closed.add(onClosed)
            
            # Store handlers to prevent garbage collection
            if not hasattr(self, '_handlers'):
                self._handlers = []
            self._handlers.extend([onHTMLEvent, onClosed])
            
            # Store palette reference for communication
            self.current_palette = palette
            
        except Exception as ex:
            self.ui.messageBox("Error setting up palette handlers: {}".format(str(ex)))

    def _initialize_html_data(self, palette):
        """Initialize HTML with hub data using sendInfoToHTML"""
        try:
            # Use threading to avoid blocking the UI
            def load_data():
                try:
                    import time
                    # Give the HTML time to load and set up fusionJavaScriptHandler
                    time.sleep(2.0)
                    
                    # Get hubs data
                    all_hubs = self.data.dataHubs
                    if all_hubs.count == 0:
                        palette.sendInfoToHTML('onError', 'No hubs found in your Fusion 360 account')
                        return
                    
                    hubs_data = []
                    for i in range(all_hubs.count):
                        hub = all_hubs.item(i)
                        hubs_data.append({
                            'id': hub.id,
                            'name': hub.name,
                            'index': i
                        })
                    
                    self.current_hubs = all_hubs
                    
                    # Send hubs data to HTML
                    hubs_json = json.dumps(hubs_data)
                    self.log.info("Sending {} hubs to HTML".format(len(hubs_data)))
                    palette.sendInfoToHTML('onHubsLoaded', hubs_json)
                    
                except Exception as ex:
                    self.log.error("Error in load_data: {}".format(str(ex)))
                    palette.sendInfoToHTML('onError', 'Error loading hubs: {}'.format(str(ex)))
            
            # Start the data loading in a separate thread
            data_thread = threading.Thread(target=load_data)
            data_thread.daemon = True
            data_thread.start()
            
        except Exception as ex:
            self.ui.messageBox("Error initializing HTML data: {}".format(str(ex)))

    def _simple_folder_selection(self):
        """Quick selection using first available hub/project"""
        try:
            # Get first hub
            all_hubs = self.data.dataHubs
            if all_hubs.count == 0:
                self.ui.messageBox("No hubs found in your Fusion 360 account.")
                return None
            
            selected_hub = all_hubs.item(0)
            
            # Get first project
            all_projects = selected_hub.dataProjects
            if all_projects.count == 0:
                self.ui.messageBox("No projects found in hub: {}".format(selected_hub.name))
                return None
            
            selected_project = all_projects.item(0)
            
            # Show what was auto-selected
            confirm = self.ui.messageBox(
                "Auto-selected:\nHub: {}\nProject: {}\n\nClick 'Yes' to continue\nClick 'No' to cancel".format(
                    selected_hub.name, selected_project.name
                ),
                "Confirm Auto-Selection",
                adsk.core.MessageBoxButtonTypes.YesNoButtonType
            )
            
            if confirm != adsk.core.DialogResults.DialogYes:
                return None
            
            # Now browse folders in this project
            return self._browse_project_folders_enhanced(selected_project)
            
        except Exception as ex:
            self.ui.messageBox("Error in simple folder selection: {}".format(str(ex)))
            return None

    def _detailed_folder_selection(self):
        """Detailed selection allowing user to choose hub/project"""
        try:
            # Hub selection
            all_hubs = self.data.dataHubs
            if all_hubs.count == 0:
                self.ui.messageBox("No hubs found in your Fusion 360 account.")
                return None
            
            selected_hub = self._select_hub(all_hubs)
            if selected_hub is None:
                return None
            
            # Project selection
            all_projects = selected_hub.dataProjects
            if all_projects.count == 0:
                self.ui.messageBox("No projects found in hub: {}".format(selected_hub.name))
                return None
            
            selected_project = self._select_project(all_projects)
            if selected_project is None:
                return None
            
            # Folder selection
            return self._browse_project_folders_enhanced(selected_project)
            
        except Exception as ex:
            self.ui.messageBox("Error in detailed folder selection: {}".format(str(ex)))
            return None

    def _select_hub(self, all_hubs):
        """Select hub with better UI"""
        if all_hubs.count == 1:
            return all_hubs.item(0)
        
        # Create a simple list for hub selection
        hub_list = []
        for i in range(min(all_hubs.count, 5)):  # Limit to first 5 hubs
            hub_list.append(all_hubs.item(i))
        
        if len(hub_list) <= 2:
            choice = self.ui.messageBox(
                "Select Hub:\n\n1. {}\n2. {}\n\nClick 'Yes' for hub 1\nClick 'No' for hub 2\nClick 'Cancel' to exit".format(
                    hub_list[0].name, hub_list[1].name if len(hub_list) > 1 else "N/A"
                ),
                "Select Hub",
                adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
            )
            if choice == adsk.core.DialogResults.DialogYes:
                return hub_list[0]
            elif choice == adsk.core.DialogResults.DialogNo and len(hub_list) > 1:
                return hub_list[1]
            else:
                return None
        else:
            # For more hubs, just use first one with confirmation
            choice = self.ui.messageBox(
                "Found {} hubs. Using first hub: {}\n\nClick 'Yes' to continue\nClick 'No' to cancel".format(
                    all_hubs.count, hub_list[0].name
                ),
                "Hub Selection",
                adsk.core.MessageBoxButtonTypes.YesNoButtonType
            )
            return hub_list[0] if choice == adsk.core.DialogResults.DialogYes else None

    def _select_project(self, all_projects):
        """Select project with better UI"""
        if all_projects.count == 1:
            return all_projects.item(0)
        
        # Create a simple list for project selection
        project_list = []
        for i in range(min(all_projects.count, 5)):  # Limit to first 5 projects
            project_list.append(all_projects.item(i))
        
        if len(project_list) <= 2:
            choice = self.ui.messageBox(
                "Select Project:\n\n1. {}\n2. {}\n\nClick 'Yes' for project 1\nClick 'No' for project 2\nClick 'Cancel' to exit".format(
                    project_list[0].name, project_list[1].name if len(project_list) > 1 else "N/A"
                ),
                "Select Project",
                adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
            )
            if choice == adsk.core.DialogResults.DialogYes:
                return project_list[0]
            elif choice == adsk.core.DialogResults.DialogNo and len(project_list) > 1:
                return project_list[1]
            else:
                return None
        else:
            # For more projects, just use first one with confirmation
            choice = self.ui.messageBox(
                "Found {} projects. Using first project: {}\n\nClick 'Yes' to continue\nClick 'No' to cancel".format(
                    all_projects.count, project_list[0].name
                ),
                "Project Selection",
                adsk.core.MessageBoxButtonTypes.YesNoButtonType
            )
            return project_list[0] if choice == adsk.core.DialogResults.DialogYes else None

    def _browse_project_folders_enhanced(self, project):
        """Enhanced folder browsing with better UX"""
        try:
            root_folder = project.rootFolder
            
            # Check if root folder has subfolders
            subfolders = []
            for folder in root_folder.dataFolders:
                subfolders.append(folder)
            
            if len(subfolders) == 0:
                # No subfolders, use root folder
                confirm_root = self.ui.messageBox(
                    "Project '{}' has no subfolders.\n\nClick 'Yes' to use root folder\nClick 'No' to cancel".format(project.name),
                    "Use Root Folder?",
                    adsk.core.MessageBoxButtonTypes.YesNoButtonType
                )
                if confirm_root == adsk.core.DialogResults.DialogYes:
                    return root_folder
                else:
                    return None
            else:
                # Build list of all folders (root + subfolders)
                all_folders = [root_folder] + subfolders
                folder_names = ["Root Folder"] + [f.name for f in subfolders]
                
                # Use enhanced folder selection
                return self._enhanced_folder_selection(project.name, all_folders, folder_names)
                    
        except Exception as ex:
            self.ui.messageBox("Error browsing project folders: {}".format(str(ex)))
            return None

    def _enhanced_folder_selection(self, project_name, folders, folder_names):
        """Enhanced folder selection with numbered options and direct choice"""
        try:
            if len(folders) <= 1:
                return folders[0] if folders else None
            
            # Show all folders with numbers
            folder_list = "\n".join("{}. {}".format(i+1, name) for i, name in enumerate(folder_names))
            
            # For 2-3 folders, allow direct selection
            if len(folders) <= 3:
                if len(folders) == 2:
                    choice = self.ui.messageBox(
                        "Project '{}' folders:\n\n{}\n\nClick 'Yes' for folder 1\nClick 'No' for folder 2\nClick 'Cancel' to exit".format(
                            project_name, folder_list
                        ),
                        "Select Folder",
                        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
                    )
                    if choice == adsk.core.DialogResults.DialogYes:
                        return folders[0]
                    elif choice == adsk.core.DialogResults.DialogNo:
                        return folders[1]
                    else:
                        return None
                else:  # 3 folders
                    choice = self.ui.messageBox(
                        "Project '{}' folders:\n\n{}\n\nClick 'Yes' for folder 1\nClick 'No' for folders 2-3\nClick 'Cancel' to exit".format(
                            project_name, folder_list
                        ),
                        "Select Folder",
                        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
                    )
                    if choice == adsk.core.DialogResults.DialogYes:
                        return folders[0]
                    elif choice == adsk.core.DialogResults.DialogNo:
                        # Choose between folders 2 and 3
                        choice2 = self.ui.messageBox(
                            "Select folder:\n\n2. {}\n3. {}\n\nClick 'Yes' for folder 2\nClick 'No' for folder 3".format(
                                folder_names[1], folder_names[2]
                            ),
                            "Select Folder 2 or 3",
                            adsk.core.MessageBoxButtonTypes.YesNoButtonType
                        )
                        return folders[1] if choice2 == adsk.core.DialogResults.DialogYes else folders[2]
                    else:
                        return None
            
            # For 4+ folders, use a text-based input approach
            return self._text_based_folder_selection(project_name, folders, folder_names)
                    
        except Exception as ex:
            self.ui.messageBox("Error in enhanced folder selection: {}".format(str(ex)))
            return None

    def _text_based_folder_selection(self, project_name, folders, folder_names):
        """Allow user to type folder number for large lists"""
        try:
            folder_list = "\n".join("{}. {}".format(i+1, name) for i, name in enumerate(folder_names))
            
            # Show the list and ask for input via a series of dialogs
            overview = self.ui.messageBox(
                "Project '{}' has {} folders:\n\n{}\n\nRemember the number of your desired folder.\nClick 'Yes' to continue with selection\nClick 'No' to cancel".format(
                    project_name, len(folders), folder_list
                ),
                "Folder List",
                adsk.core.MessageBoxButtonTypes.YesNoButtonType
            )
            
            if overview != adsk.core.DialogResults.DialogYes:
                return None
            
            # Use a binary search approach to find the folder
            return self._binary_folder_search(folders, folder_names, 1, len(folders))
                    
        except Exception as ex:
            self.ui.messageBox("Error in text-based folder selection: {}".format(str(ex)))
            return None

    def _binary_folder_search(self, folders, folder_names, start_num, end_num):
        """Use binary search approach to narrow down folder selection"""
        try:
            if start_num == end_num:
                return folders[start_num - 1]
            
            if end_num - start_num == 1:
                # Two options left
                choice = self.ui.messageBox(
                    "Select folder:\n\n{}. {}\n{}. {}\n\nClick 'Yes' for folder {}\nClick 'No' for folder {}".format(
                        start_num, folder_names[start_num - 1],
                        end_num, folder_names[end_num - 1],
                        start_num, end_num
                    ),
                    "Final Selection",
                    adsk.core.MessageBoxButtonTypes.YesNoButtonType
                )
                return folders[start_num - 1] if choice == adsk.core.DialogResults.DialogYes else folders[end_num - 1]
            
            # Split the range
            mid_point = (start_num + end_num) // 2
            
            choice = self.ui.messageBox(
                "Is your folder number {} or lower?\n\nFolders {}-{}:\n{}\n\nFolders {}-{}:\n{}\n\nClick 'Yes' for folders {}-{}\nClick 'No' for folders {}-{}\nClick 'Cancel' to exit".format(
                    mid_point,
                    start_num, mid_point,
                    "\n".join("{}. {}".format(i, folder_names[i-1]) for i in range(start_num, mid_point + 1)),
                    mid_point + 1, end_num,
                    "\n".join("{}. {}".format(i, folder_names[i-1]) for i in range(mid_point + 1, end_num + 1)),
                    start_num, mid_point,
                    mid_point + 1, end_num
                ),
                "Narrow Down Selection",
                adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
            )
            
            if choice == adsk.core.DialogResults.DialogYes:
                return self._binary_folder_search(folders, folder_names, start_num, mid_point)
            elif choice == adsk.core.DialogResults.DialogNo:
                return self._binary_folder_search(folders, folder_names, mid_point + 1, end_num)
            else:
                return None
                
        except Exception as ex:
            self.ui.messageBox("Error in binary folder search: {}".format(str(ex)))
            return None

    def _browse_project_folders(self, project):
        """Legacy method - redirect to enhanced version"""
        return self._browse_project_folders_enhanced(project)

    def _select_folder_from_list(self, project_name, folders, folder_names):
        """Let user select a folder from a numbered list"""
        try:
            if len(folders) <= 1:
                return folders[0] if folders else None
            
            # For 2 folders, use simple Yes/No selection
            if len(folders) == 2:
                choice = self.ui.messageBox(
                    "Project '{}' has 2 folders:\n\n1. {}\n2. {}\n\nClick 'Yes' for folder 1\nClick 'No' for folder 2\nClick 'Cancel' to exit".format(
                        project_name, folder_names[0], folder_names[1]
                    ),
                    "Select Folder",
                    adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
                )
                if choice == adsk.core.DialogResults.DialogYes:
                    return folders[0]
                elif choice == adsk.core.DialogResults.DialogNo:
                    return folders[1]
                else:
                    return None
            
            # For 3+ folders, use grouped selection approach
            return self._select_from_multiple_folders(project_name, folders, folder_names)
                    
        except Exception as ex:
            self.ui.messageBox("Error in folder selection: {}".format(str(ex)))
            return None

    def _select_from_multiple_folders(self, project_name, folders, folder_names):
        """Handle selection from 3+ folders using grouped approach"""
        try:
            # Group folders into sets of 3 for easier selection
            total_folders = len(folders)
            
            # Show all folders first
            folder_list = "\n".join("{}. {}".format(i+1, name) for i, name in enumerate(folder_names))
            overview = self.ui.messageBox(
                "Project '{}' has {} folders:\n\n{}\n\nClick 'Yes' to select by number\nClick 'No' to cancel".format(
                    project_name, total_folders, folder_list
                ),
                "Available Folders",
                adsk.core.MessageBoxButtonTypes.YesNoButtonType
            )
            
            if overview != adsk.core.DialogResults.DialogYes:
                return None
            
            # For folders 1-3, use direct selection
            if total_folders <= 3:
                if total_folders == 3:
                    choice = self.ui.messageBox(
                        "Select folder:\n\n1. {}\n2. {}\n3. {}\n\nClick 'Yes' for folder 1\nClick 'No' for folders 2-3\nClick 'Cancel' to exit".format(
                            folder_names[0], folder_names[1], folder_names[2]
                        ),
                        "Select Folder",
                        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
                    )
                    if choice == adsk.core.DialogResults.DialogYes:
                        return folders[0]
                    elif choice == adsk.core.DialogResults.DialogNo:
                        # Choose between folders 2 and 3
                        choice2 = self.ui.messageBox(
                            "Select folder:\n\n2. {}\n3. {}\n\nClick 'Yes' for folder 2\nClick 'No' for folder 3".format(
                                folder_names[1], folder_names[2]
                            ),
                            "Select Folder 2 or 3",
                            adsk.core.MessageBoxButtonTypes.YesNoButtonType
                        )
                        return folders[1] if choice2 == adsk.core.DialogResults.DialogYes else folders[2]
                    else:
                        return None
            
            # For 4+ folders, use range-based selection
            return self._select_from_many_folders(folders, folder_names)
            
        except Exception as ex:
            self.ui.messageBox("Error in multi-folder selection: {}".format(str(ex)))
            return None

    def _select_from_many_folders(self, folders, folder_names):
        """Handle selection from 4+ folders using range approach"""
        try:
            total_folders = len(folders)
            
            # Group into ranges for easier selection
            mid_point = total_folders // 2
            
            # First, choose range
            range_choice = self.ui.messageBox(
                "Select folder range:\n\nFolders 1-{}:\n{}\n\nFolders {}-{}:\n{}\n\nClick 'Yes' for folders 1-{}\nClick 'No' for folders {}-{}\nClick 'Cancel' to exit".format(
                    mid_point,
                    "\n".join("{}. {}".format(i+1, folder_names[i]) for i in range(mid_point)),
                    mid_point + 1, total_folders,
                    "\n".join("{}. {}".format(i+1, folder_names[i]) for i in range(mid_point, total_folders)),
                    mid_point, mid_point + 1, total_folders
                ),
                "Select Folder Range",
                adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
            )
            
            if range_choice == adsk.core.DialogResults.DialogCancel:
                return None
            elif range_choice == adsk.core.DialogResults.DialogYes:
                # First half
                selected_folders = folders[:mid_point]
                selected_names = folder_names[:mid_point]
            else:
                # Second half
                selected_folders = folders[mid_point:]
                selected_names = folder_names[mid_point:]
            
            # Now select from the chosen range
            if len(selected_folders) == 1:
                return selected_folders[0]
            else:
                return self._select_folder_from_list("Range", selected_folders, selected_names)
                
        except Exception as ex:
            self.ui.messageBox("Error in range-based selection: {}".format(str(ex)))
            return None

    def _export_2d_data(self, output_path):
        progress_dialog = self.ui.createProgressDialog()
        progress_dialog.show("Exporting 2D drawings!", "", 0, 1, 1)

        # Export from selected cloud folder and all its sub-folders
        self.log.info("Scanning cloud folder and all sub-folders: {}".format(self.selected_cloud_folder.name))
        
        # Find all .f3d and .f3z files in the cloud folder and all sub-folders recursively
        drawing_files = self._get_files_for(self.selected_cloud_folder)
        
        if not drawing_files:
            self.ui.messageBox("No Fusion 360 files (.f3d or .f3z) found in the selected cloud folder or its sub-folders.")
            return
        
        progress_dialog.message = "Exporting {} drawing files from folder and sub-folders\nProcessing file %v of %m".format(len(drawing_files))
        progress_dialog.maximumValue = len(drawing_files)
        progress_dialog.reset()
        
        self.log.info("Found {} Fusion 360 files to process (including sub-folders)".format(len(drawing_files)))
        
        for file_index, data_file in enumerate(drawing_files):
            if progress_dialog.wasCancelled:
                self.log.info("The process was cancelled!")
                self.was_cancelled = True
                return

            progress_dialog.progressValue = file_index + 1
            self._write_2d_data_file(output_path, data_file)
        
        self.log.info("Finished exporting from cloud folder")

    def _find_drawing_files_in_folder(self, folder_path):
        """Find all Fusion 360 files in the specified folder"""
        drawing_files = []
        
        try:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith(('.f3d', '.f3z')):
                        full_path = os.path.join(root, file)
                        drawing_files.append(full_path)
                        
        except Exception as ex:
            self.log.error("Error scanning folder {}: {}".format(folder_path, str(ex)))
            
        return drawing_files

    def _process_drawing_file(self, output_path, file_path):
        """Process a single drawing file from the file system"""
        file_name = os.path.basename(file_path)
        self.log.info("Processing file: {}".format(file_name))
        
        try:
            # Open the document from file path
            document = self.app.documents.open(file_path)
            
            if document is None:
                raise Exception("Failed to open file: {}".format(file_path))
                
            document.activate()
            
            # Create output folder for this file
            base_name = os.path.splitext(file_name)[0]
            file_output_path = self._take(output_path, self._name(base_name))
            
            self.log.info("Exporting to: {}".format(file_output_path))
            
            # Get the design and export 2D content
            fusion_document = adsk.fusion.FusionDocument.cast(document)
            design = fusion_document.design
            
            # Export only drawings (blueprints with dimensions)
            if 'drawings' in self.file_types_to_export:
                self._export_drawings(file_output_path, design)
            
            self.log.info("Finished processing: {}".format(file_name))
            
        except Exception as ex:
            self.num_issues += 1
            self.log.exception("Failed to process file {}: {}".format(file_name, str(ex)))
            
        finally:
            try:
                if 'document' in locals() and document is not None:
                    document.close(False)
                    
                # Clear cache if enabled
                if self.clear_cache_after_each_file:
                    self._clear_fusion_cache()
                    
            except Exception as ex:
                self.num_issues += 1
                self.log.exception("Failed to close file {}: {}".format(file_name, str(ex)))

    def _export_selected_folders(self, output_path, progress_dialog):
        """Export files from only the selected folders"""
        all_files = []
        
        # Collect files from selected folders
        for folder in self.selected_folders:
            self.log.info("Processing selected folder \"{}\"".format(folder.name))
            files = self._get_files_for(folder)
            all_files.extend(files)
        
        progress_dialog.message = "Exporting from {} selected folder{}\nExporting design %v of %m".format(
            len(self.selected_folders),
            "s" if len(self.selected_folders) > 1 else ""
        )
        progress_dialog.maximumValue = len(all_files)
        progress_dialog.reset()
        
        if not all_files:
            self.log.info("No files found in selected folders")
            return
        
        self.log.info("Found {} files in selected folders".format(len(all_files)))
        
        for file_index in range(len(all_files)):
            if progress_dialog.wasCancelled:
                self.log.info("The process was cancelled!")
                self.was_cancelled = True
                return

            file: adsk.core.DataFile = all_files[file_index]
            progress_dialog.progressValue = file_index + 1
            self._write_2d_data_file(output_path, file)
        
        self.log.info("Finished exporting from selected folders")

    def _ask_for_output_path(self):
        folder_dialog = self.ui.createFolderDialog()
        folder_dialog.title = "Where should we store the 2D export?"
        dialog_result = folder_dialog.showDialog()
        if dialog_result != adsk.core.DialogResults.DialogOK:
            return None

        output_path = folder_dialog.folder
        return output_path

    def _get_files_for(self, folder):
        """Recursively get all files from a folder and all its sub-folders"""
        files = []
        
        # Get files directly in this folder
        for file in folder.dataFiles:
            files.append(file)
        
        # Log folder processing
        if folder.dataFiles.count > 0:
            self.log.info("Found {} files in folder: {}".format(folder.dataFiles.count, folder.name))
        
        # Recursively process all sub-folders
        sub_folder_count = folder.dataFolders.count
        if sub_folder_count > 0:
            self.log.info("Processing {} sub-folders in: {}".format(sub_folder_count, folder.name))
            
            for sub_folder in folder.dataFolders:
                sub_files = self._get_files_for(sub_folder)
                files.extend(sub_files)
                if len(sub_files) > 0:
                    self.log.info("Added {} files from sub-folder: {}".format(len(sub_files), sub_folder.name))

        return files

    def _write_2d_data_file(self, root_folder, file: adsk.core.DataFile):
        if file.fileExtension != "f3d" and file.fileExtension != "f3z":
            self.log.info("Skipping non-Fusion file \"{}\"".format(file.name))
            return

        self.log.info("Processing file \"{}\" for 2D export".format(file.name))

        try:
            document = self.documents.open(file)

            if document is None:
                raise Exception("Documents.open returned None")

            document.activate()
        except BaseException as ex:
            self.num_issues += 1
            self.log.exception("Opening {} failed!".format(file.name), exc_info=ex)
            return

        try:
            file_folder = file.parentFolder
            file_folder_path = self._name(file_folder.name)

            while file_folder.parentFolder is not None:
                file_folder = file_folder.parentFolder
                file_folder_path = os.path.join(self._name(file_folder.name), file_folder_path)

            parent_project = file_folder.parentProject
            parent_hub = parent_project.parentHub

            file_folder_path = self._take(
                root_folder,
                "Hub {}".format(self._name(parent_hub.name)),
                "Project {}".format(self._name(parent_project.name)),
                file_folder_path,
                self._name(file.name)
            )

            if not os.path.exists(file_folder_path):
                self.num_issues += 1
                self.log.exception("Couldn't make root folder\"{}\"".format(file_folder_path))
                return

            self.log.info("Writing 2D data to \"{}\"".format(file_folder_path))

            fusion_document: adsk.fusion.FusionDocument = adsk.fusion.FusionDocument.cast(document)
            design: adsk.fusion.Design = fusion_document.design

            # Export 2D data based on selected types
            if 'sketches' in self.file_types_to_export:
                self._export_sketches(file_folder_path, design.rootComponent)
                
            if 'drawings' in self.file_types_to_export:
                self._export_drawings(file_folder_path, design)

            self.log.info("Finished exporting 2D data from file \"{}\"".format(file.name))
            
        except BaseException as ex:
            self.num_issues += 1
            self.log.exception("Failed while working on \"{}\"".format(file.name), exc_info=ex)
        finally:
            try:
                if document is not None:
                    document.close(False)
                    
                # Clear cache if option is enabled
                if self.clear_cache_after_each_file:
                    self._clear_fusion_cache()
                    
            except BaseException as ex:
                self.num_issues += 1
                self.log.exception("Failed to close \"{}\"".format(file.name), exc_info=ex)

    def _export_sketches(self, base_path, component: adsk.fusion.Component):
        """Export all sketches from a component and its sub-components"""
        sketches = component.sketches
        
        if sketches.count > 0:
            sketches_path = self._take(base_path, "Sketches")
            
            for sketch_index in range(sketches.count):
                sketch = sketches.item(sketch_index)
                sketch_name = self._name(sketch.name) if sketch.name else f"Sketch_{sketch_index + 1}"
                
                for format_type in self.export_formats:
                    self._export_sketch_in_format(sketches_path, sketch, sketch_name, format_type)

        # Process sub-components recursively
        occurrences = component.occurrences
        for occurrence_index in range(occurrences.count):
            occurrence = occurrences.item(occurrence_index)
            sub_component = occurrence.component
            self._export_sketches(base_path, sub_component)

    def _export_drawings(self, base_path, design: adsk.fusion.Design):
        """Export all drawings from the design"""
        drawings = design.drawings
        
        if drawings.count > 0:
            drawings_path = self._take(base_path, "Drawings")
            
            for drawing_index in range(drawings.count):
                drawing = drawings.item(drawing_index)
                drawing_name = self._name(drawing.name) if drawing.name else f"Drawing_{drawing_index + 1}"
                
                for format_type in self.export_formats:
                    self._export_drawing_in_format(drawings_path, drawing, drawing_name, format_type)

    def _export_sketch_in_format(self, output_path, sketch: adsk.fusion.Sketch, sketch_name: str, format_type: str):
        """Export a sketch in the specified format"""
        file_path = os.path.join(output_path, f"{sketch_name}.{format_type}")
        
        if os.path.exists(file_path):
            self.log.info("Sketch file \"{}\" already exists".format(file_path))
            return

        self.log.info("Writing sketch file \"{}\"".format(file_path))

        try:
            if format_type == 'dxf':
                sketch.saveAsDXF(file_path)
            elif format_type == 'dwg':
                # Note: DWG export might not be available in all Fusion 360 versions
                sketch.saveAsDXF(file_path.replace('.dwg', '.dxf'))
                self.log.warning("DWG format not directly supported, exported as DXF instead")
            else:
                self.log.warning("Format {} not supported for sketches, skipping".format(format_type))
                
        except BaseException as ex:
            self.num_issues += 1
            self.log.exception("Failed to export sketch \"{}\" in format {}".format(sketch_name, format_type), exc_info=ex)

    def _export_drawing_in_format(self, output_path, drawing, drawing_name: str, format_type: str):
        """Export a drawing in the specified format"""
        file_path = os.path.join(output_path, f"{drawing_name}.{format_type}")
        
        if os.path.exists(file_path):
            self.log.info("Drawing file \"{}\" already exists".format(file_path))
            return

        self.log.info("Writing drawing file \"{}\"".format(file_path))

        try:
            # Get the active viewport for the drawing
            if drawing.sheets.count > 0:
                sheet = drawing.sheets.item(0)  # Use first sheet
                
                if format_type == 'pdf':
                    drawing.saveAsPDF(file_path)
                elif format_type == 'dxf':
                    sheet.saveAsDXF(file_path)
                elif format_type == 'dwg':
                    # Note: DWG export might not be available in all versions
                    sheet.saveAsDXF(file_path.replace('.dwg', '.dxf'))
                    self.log.warning("DWG format not directly supported, exported as DXF instead")
                elif format_type in ['png', 'jpg', 'svg']:
                    # For image formats, we would need to use the viewport export
                    self.log.warning("Format {} not yet implemented for drawings".format(format_type))
                else:
                    self.log.warning("Format {} not supported for drawings".format(format_type))
            else:
                self.log.warning("Drawing \"{}\" has no sheets to export".format(drawing_name))
                
        except BaseException as ex:
            self.num_issues += 1
            self.log.exception("Failed to export drawing \"{}\" in format {}".format(drawing_name, format_type), exc_info=ex)

    def _take(self, *path):
        out_path = os.path.join(*path)
        os.makedirs(out_path, exist_ok=True)
        return out_path

    def _name(self, name):
        """Sanitize filename by removing invalid characters"""
        if not name:
            return "Unnamed"
            
        name = re.sub('[^a-zA-Z0-9 \n\.]', '_', name).strip()
        
        # Handle common file extensions that might conflict
        if name.endswith('.dxf') or name.endswith('.dwg') or name.endswith('.pdf'):
            name = name[0: -4] + "_" + name[-3:]

        return name

    def _clear_fusion_cache(self):
        """Clear Fusion 360 cache folders to free up memory"""
        try:
            cache_folders = []
            
            # Windows cache locations
            if os.name == 'nt':
                localappdata = os.environ.get('LOCALAPPDATA', '')
                temp_dir = tempfile.gettempdir()
                
                cache_folders.extend([
                    os.path.join(localappdata, 'Autodesk', 'Webdeploy', 'Production'),
                    os.path.join(temp_dir, 'Autodesk', 'Fusion360'),
                    os.path.join(localappdata, 'Autodesk', 'Fusion 360', 'Cache'),
                    os.path.join(localappdata, 'Autodesk', 'Fusion 360', 'Temp')
                ])
            
            # macOS cache locations
            else:
                home_dir = os.path.expanduser('~')
                temp_dir = tempfile.gettempdir()
                
                cache_folders.extend([
                    os.path.join(home_dir, 'Library', 'Application Support', 'Autodesk', 'Fusion 360', 'Cache'),
                    os.path.join(temp_dir, 'Autodesk', 'Fusion360'),
                    os.path.join(home_dir, 'Library', 'Caches', 'com.autodesk.fusion360')
                ])
            
            cleared_count = 0
            for cache_folder in cache_folders:
                if os.path.exists(cache_folder):
                    try:
                        # Clear contents but keep the folder structure
                        for item in os.listdir(cache_folder):
                            item_path = os.path.join(cache_folder, item)
                            if os.path.isfile(item_path):
                                os.remove(item_path)
                                cleared_count += 1
                            elif os.path.isdir(item_path):
                                shutil.rmtree(item_path, ignore_errors=True)
                                cleared_count += 1
                    except (OSError, PermissionError) as ex:
                        # Some cache files might be locked, continue with others
                        self.log.warning("Could not clear cache folder {}: {}".format(cache_folder, str(ex)))
                        continue
            
            if cleared_count > 0:
                self.log.info("Cleared {} cache items".format(cleared_count))
            
        except BaseException as ex:
            self.log.warning("Cache clearing failed: {}".format(str(ex)))


# Event handler classes for HTML palette communication
class HTMLEventHandler(adsk.core.HTMLEventHandler):
    def __init__(self, exporter):
        super().__init__()
        self.exporter = exporter

    def notify(self, args):
        try:
            htmlArgs = adsk.core.HTMLEventArgs.cast(args)
            action = htmlArgs.action
            data = htmlArgs.data

            self.exporter.log.info("HTML Event: action='{}', data='{}'".format(action, data))

            # Get the palette for sendInfoToHTML calls
            palette = None
            if hasattr(self.exporter, 'current_palette'):
                palette = self.exporter.current_palette
            
            if action == 'getProjects':
                hub_index = int(data)
                self._send_projects_to_html(palette, hub_index)
            elif action == 'getFolders':
                project_index = int(data)
                self._send_folders_to_html(palette, project_index)
            elif action == 'selectFolder':
                self._handle_folder_selection(data)
            elif action == 'cancel':
                self.exporter.selected_folder_result = None
                self.exporter.palette_closed = True

        except Exception as ex:
            self.exporter.log.error("HTML Event Error: {}".format(str(ex)))
            self.exporter.ui.messageBox("HTML Event Error: {}".format(str(ex)))

    def _send_hubs_to_html(self, htmlArgs):
        try:
            all_hubs = self.exporter.data.dataHubs
            hubs_data = []
            
            for i in range(all_hubs.count):
                hub = all_hubs.item(i)
                hubs_data.append({
                    'id': hub.id,
                    'name': hub.name,
                    'index': i
                })
            
            self.exporter.current_hubs = all_hubs
            htmlArgs.returnData = json.dumps(hubs_data)
            
        except Exception as ex:
            htmlArgs.returnData = json.dumps([])
            self.exporter.ui.messageBox("Error loading hubs: {}".format(str(ex)))

    def _send_projects_to_html(self, palette, hub_index):
        try:
            if not hasattr(self.exporter, 'current_hubs') or hub_index >= self.exporter.current_hubs.count:
                if palette:
                    palette.sendInfoToHTML('onProjectsLoaded', json.dumps([]))
                return
                
            hub = self.exporter.current_hubs.item(hub_index)
            all_projects = hub.dataProjects
            projects_data = []
            
            for i in range(all_projects.count):
                project = all_projects.item(i)
                projects_data.append({
                    'id': project.id,
                    'name': project.name,
                    'index': i
                })
            
            self.exporter.current_projects = all_projects
            if palette:
                palette.sendInfoToHTML('onProjectsLoaded', json.dumps(projects_data))
            
        except Exception as ex:
            if palette:
                palette.sendInfoToHTML('onProjectsLoaded', json.dumps([]))
            self.exporter.log.error("Error loading projects: {}".format(str(ex)))

    def _send_folders_to_html(self, palette, project_index):
        try:
            if not hasattr(self.exporter, 'current_projects') or project_index >= self.exporter.current_projects.count:
                if palette:
                    palette.sendInfoToHTML('onFoldersLoaded', json.dumps([]))
                return
                
            project = self.exporter.current_projects.item(project_index)
            root_folder = project.rootFolder
            folders_data = []
            
            # Add root folder
            folders_data.append({
                'id': root_folder.id,
                'name': 'Root Folder',
                'isRoot': True
            })
            
            # Add subfolders
            for folder in root_folder.dataFolders:
                folders_data.append({
                    'id': folder.id,
                    'name': folder.name,
                    'isRoot': False
                })
            
            self.exporter.current_folders = {
                'root': root_folder,
                'all': folders_data
            }
            if palette:
                palette.sendInfoToHTML('onFoldersLoaded', json.dumps(folders_data))
            
        except Exception as ex:
            if palette:
                palette.sendInfoToHTML('onFoldersLoaded', json.dumps([]))
            self.exporter.log.error("Error loading folders: {}".format(str(ex)))

    def _handle_folder_selection(self, folder_id):
        try:
            if not hasattr(self.exporter, 'current_folders'):
                return
                
            # Find the selected folder
            if folder_id == self.exporter.current_folders['root'].id:
                self.exporter.selected_folder_result = self.exporter.current_folders['root']
            else:
                # Search in subfolders
                root_folder = self.exporter.current_folders['root']
                for folder in root_folder.dataFolders:
                    if folder.id == folder_id:
                        self.exporter.selected_folder_result = folder
                        break
            
            self.exporter.palette_closed = True
            
        except Exception as ex:
            self.exporter.ui.messageBox("Error selecting folder: {}".format(str(ex)))



class PaletteClosedHandler(adsk.core.UserInterfaceGeneralEventHandler):
    def __init__(self, exporter):
        super().__init__()
        self.exporter = exporter

    def notify(self, args):
        self.exporter.palette_closed = True


def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()

        with TwoDExport(app) as two_d_export:
            two_d_export.run(context)

    except:
        ui = app.userInterface
        ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
