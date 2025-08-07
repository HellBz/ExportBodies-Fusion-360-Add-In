import adsk.core
import os
import traceback
import re
from ... import config
from ...lib import fusionAddInUtils as futil

# Get application and UI instances
app = adsk.core.Application.get()
ui = app.userInterface

# Command metadata
CMD_NAME = 'Export all Bodies' #os.path.basename(os.path.dirname(__file__))
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_{CMD_NAME}'
CMD_Description = 'Export all bodies to selected formats'
IS_PROMOTED = True

# Workspace and panel configuration
WORKSPACE_ID = config.design_workspace
TAB_ID = config.tools_tab_id
TAB_NAME = config.my_tab_name
PANEL_ID = config.my_panel_id
PANEL_NAME = config.my_panel_name
PANEL_AFTER = config.my_panel_after

# Icon directory
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Event handler storage
local_handlers = []

# Called when the add-in starts
def start():
    # Create the button definition
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)
    futil.add_handler(cmd_def.commandCreated, command_created)

    # Get or create the toolbar tab and panel
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID)
    if toolbar_tab is None:
        toolbar_tab = workspace.toolbarTabs.add(TAB_ID, TAB_NAME)

    panel = toolbar_tab.toolbarPanels.itemById(PANEL_ID)
    if panel is None:
        panel = toolbar_tab.toolbarPanels.add(PANEL_ID, PANEL_NAME, PANEL_AFTER, False)

    # Add the button to the panel
    control = panel.controls.addCommand(cmd_def)
    control.isPromoted = IS_PROMOTED

# Called when the add-in is stopped
def stop():
    # Clean up the created controls and definitions
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    if command_control:
        command_control.deleteMe()
    if command_definition:
        command_definition.deleteMe()
    if panel.controls.count == 0:
        panel.deleteMe()
    if toolbar_tab.toolbarPanels.count == 0:
        toolbar_tab.deleteMe()

# Called when the command is created
# Sets up the dialog UI and options
def command_created(args: adsk.core.CommandCreatedEventArgs):
    futil.log(f'{CMD_NAME} Command Created Event')

    # Attach handlers
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    # Setup command input UI
    inputs = args.command.commandInputs
    group = inputs.addGroupCommandInput('formatGroup', 'Export Formats')
    format_inputs = group.children

    # Load last used config (export formats)
    this_dir = os.path.dirname(os.path.realpath(__file__))
    config_file = os.path.join(this_dir, 'last_export_config.txt')
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            last = [x.strip().upper() for x in f.read().strip().split(',')]
    else:
        last = ['STL', '3MF']

    # Available formats
    formats = ['STL', '3MF', 'STEP', 'IGES', 'SAT', 'SMT', 'F3D']
    for fmt in formats:
        checked = fmt in last
        format_inputs.addBoolValueInput(fmt, fmt, True, '', checked)

# Called when the command is executed
# Handles format selection and triggers the export process
def command_execute(args: adsk.core.CommandEventArgs):
    futil.log(f'{CMD_NAME} Command Execute Event')

    # Read selected formats
    inputs = args.command.commandInputs
    format_group = inputs.itemById('formatGroup')
    format_inputs = [format_group.children.item(i) for i in range(format_group.children.count)]
    selected_formats = [f.name for f in format_inputs if f.value]

    if not selected_formats:
        ui.messageBox('No formats selected.')
        return

    # Save the last used config
    try:
        this_dir = os.path.dirname(os.path.realpath(__file__))
        config_file = os.path.join(this_dir, 'last_export_config.txt')
        with open(config_file, 'w') as f:
            f.write(','.join(selected_formats))
    except:
        ui.messageBox('Could not save config')

    do_export(selected_formats)

# Performs the actual export process for each body and format

def do_export(selected_formats):
    design = app.activeProduct
    exportMgr = design.exportManager
    rootComp = design.rootComponent
    bodies = rootComp.bRepBodies

    if bodies.count == 0:
        ui.messageBox('No bodies found to export.')
        return

    # Build output path
    doc_name = app.activeDocument.name
    base_name, ext = os.path.splitext(doc_name)

    version_match = re.search(r'(v\d+)', base_name, re.IGNORECASE)
    if version_match:
        version = version_match.group(1)
        clean_base = base_name[:version_match.start()].strip()
    else:
        version = ''
        clean_base = base_name

    safe_base_name = clean_base.replace(" ", "_").replace(".", "_")
    safe_version = version.replace(" ", "_") if version else ''

    # Get the absolute path of this script file
    this_dir = os.path.dirname(os.path.realpath(__file__))

    # Go two levels up from this directory
    parent_dir = os.path.dirname(os.path.dirname(this_dir))

    # Define the export base folder in that parent directory
    export_base = os.path.join(parent_dir, 'EXPORTED')
    # Create the final export folder by appending the sanitized document name
    export_folder = os.path.join(export_base, safe_base_name)

    # Ask user for export folder
    choice = ui.messageBox(
        f'Default export folder:\n{export_folder}\n\nUse this folder?',
        'Export Folder',
        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
    )
    if choice == adsk.core.DialogResults.DialogNo:
        dlg = ui.createFolderDialog()
        dlg.title = 'Select Export Folder'
        result = dlg.showDialog()
        if result == adsk.core.DialogResults.DialogOK:
            export_folder = dlg.folder
        else:
            ui.messageBox('Export canceled.')
            return

    if not os.path.exists(export_folder):
        os.makedirs(export_folder)

    # Loop over each body and export to selected formats
    for i in range(bodies.count):
        body = bodies.item(i)
        safe_body_name = body.name.replace(" ", "_")

        for fmt in selected_formats:
            filename = f'{safe_body_name}_{safe_version}.{fmt.lower()}' if safe_version else f'{safe_body_name}.{fmt.lower()}'
            full_path = os.path.join(export_folder, filename)

            # Export using correct Fusion 360 options based on format
            if fmt == 'STL':
                stlOpts = exportMgr.createSTLExportOptions(body)
                stlOpts.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementHigh
                stlOpts.filename = full_path
                exportMgr.execute(stlOpts)

            elif fmt == '3MF':
                m3fOpts = exportMgr.createC3MFExportOptions(body)
                m3fOpts.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementHigh
                m3fOpts.filename = full_path
                exportMgr.execute(m3fOpts)

            elif fmt == 'STEP':
                stepOpts = exportMgr.createSTEPExportOptions(full_path)
                exportMgr.execute(stepOpts)

            elif fmt == 'IGES':
                igesOpts = exportMgr.createIGESExportOptions(full_path)
                exportMgr.execute(igesOpts)

            elif fmt == 'SAT':
                satOpts = exportMgr.createSATExportOptions(full_path)
                exportMgr.execute(satOpts)

            elif fmt == 'SMT':
                smtOpts = exportMgr.createSMTExportOptions(full_path)
                exportMgr.execute(smtOpts)

            elif fmt == 'F3D':
                f3dOpts = exportMgr.createFusionArchiveExportOptions(full_path)
                exportMgr.execute(f3dOpts)

            else:
                ui.messageBox(f'Unknown format: {fmt}')

    # Prompt to open the export folder
    choice = ui.messageBox(
        f'Export completed:\n{export_folder}\n\nOpen folder now?',
        'Export Finished',
        adsk.core.MessageBoxButtonTypes.YesNoButtonType
    )
    if choice == adsk.core.DialogResults.DialogYes:
        try:
            os.startfile(export_folder)
        except Exception as e:
            ui.messageBox(f'Could not open folder:\n{str(e)}')

# Called when the command ends
# Cleanup event handlers

def command_destroy(args: adsk.core.CommandEventArgs):
    global local_handlers
    local_handlers = []
    futil.log(f'{CMD_NAME} Command Destroy Event')