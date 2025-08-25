import adsk.core, adsk.fusion, adsk.cam, traceback
import json

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        # Test basic hub access
        ui.messageBox("Testing hub connection...")
        
        # Get data manager
        data = app.data
        
        # Get hubs
        all_hubs = data.dataHubs
        hub_count = all_hubs.count
        
        ui.messageBox("Found {} hubs".format(hub_count))
        
        if hub_count == 0:
            ui.messageBox("No hubs found. Make sure you're logged into Fusion 360 and have access to cloud data.")
            return
        
        # List all hubs
        hub_info = []
        for i in range(hub_count):
            hub = all_hubs.item(i)
            hub_info.append("Hub {}: {} (ID: {})".format(i, hub.name, hub.id))
        
        ui.messageBox("Hubs found:\n" + "\n".join(hub_info))
        
        # Test HTML palette with simple sendInfoToHTML
        palettes = ui.palettes
        palette = palettes.itemById('testHubPalette')
        
        if palette:
            palette.deleteMe()
        
        # Create new palette
        palette = palettes.add('testHubPalette', 'Test Hub Connection', 'test_hub_palette.html', True, True, True, 400, 300)
        
        if palette:
            ui.messageBox("HTML palette created successfully")
            
            # Prepare hub data
            hubs_data = []
            for i in range(hub_count):
                hub = all_hubs.item(i)
                hubs_data.append({
                    'id': hub.id,
                    'name': hub.name,
                    'index': i
                })
            
            hubs_json = json.dumps(hubs_data)
            ui.messageBox("Sending hub data: {}".format(hubs_json[:100]))
            
            # Try sending data directly with sendInfoToHTML
            import time
            time.sleep(2)
            palette.sendInfoToHTML('onHubsLoaded', hubs_json)
            ui.messageBox("Data sent via sendInfoToHTML")
            
        else:
            ui.messageBox("Failed to create HTML palette")
            
    except Exception as ex:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
