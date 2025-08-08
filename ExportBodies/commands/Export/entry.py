import adsk.core
import os
import traceback
import re
import platform
from ... import config
from ...lib import fusionAddInUtils as futil

# Access the Fusion 360 application and UI
app = adsk.core.Application.get()
ui = app.userInterface

# Command metadata
CMD_NAME = 'Export all Bodies'
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_{CMD_NAME}'
CMD_Description = 'Export all visible or all bodies to selected formats'
IS_PROMOTED = True

# UI placement configuration
WORKSPACE_ID = config.design_workspace
TAB_ID = config.tools_tab_id
TAB_NAME = config.my_tab_name
PANEL_ID = config.my_panel_id
PANEL_NAME = config.my_panel_name
PANEL_AFTER = config.my_panel_after

# Path to command icon
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Local handler storage
local_handlers = []

# Called when the add-in is loaded
def start():
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)
    futil.add_handler(cmd_def.commandCreated, command_created)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID) or workspace.toolbarTabs.add(TAB_ID, TAB_NAME)
    panel = toolbar_tab.toolbarPanels.itemById(PANEL_ID) or toolbar_tab.toolbarPanels.add(PANEL_ID, PANEL_NAME, PANEL_AFTER, False)
    control = panel.controls.addCommand(cmd_def)
    control.isPromoted = IS_PROMOTED

# Called when the add-in is unloaded
def stop():
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID)

    if panel:
        control = panel.controls.itemById(CMD_ID)
        if control:
            control.deleteMe()

    definition = ui.commandDefinitions.itemById(CMD_ID)
    if definition:
        definition.deleteMe()

    if panel and panel.controls.count == 0:
        panel.deleteMe()
    if toolbar_tab and toolbar_tab.toolbarPanels.count == 0:
        toolbar_tab.deleteMe()

# Called when the user launches the command
def command_created(args: adsk.core.CommandCreatedEventArgs):
    futil.log(f'{CMD_NAME} Command Created Event')

    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    inputs = args.command.commandInputs
    group = inputs.addGroupCommandInput('formatGroup', 'Export Formats')
    format_inputs = group.children

    # Try loading previous export format config
    this_dir = os.path.dirname(os.path.realpath(__file__))
    config_file = os.path.join(this_dir, 'last_export_config.txt')
    last_formats = []
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            last_formats = [x.strip().upper() for x in f.read().split(',')]
    else:
        last_formats = ['STL', '3MF']

    # Add format checkboxes
    formats = ['STL', '3MF', 'STEP', 'IGES', 'SAT', 'SMT', 'F3D']
    for fmt in formats:
        format_inputs.addBoolValueInput(fmt, fmt, True, '', fmt in last_formats)

    # Checkbox: Export only visible bodies
    inputs.addBoolValueInput('onlyVisibleBodies', 'Only export visible bodies', True, '', True)

# Called when the command is executed
def command_execute(args: adsk.core.CommandEventArgs):
    futil.log(f'{CMD_NAME} Command Execute Event')

    inputs = args.command.commandInputs
    format_group = inputs.itemById('formatGroup')
    format_inputs = [format_group.children.item(i) for i in range(format_group.children.count)]
    selected_formats = [f.name for f in format_inputs if f.value]

    if not selected_formats:
        ui.messageBox('No formats selected.')
        return

    # Save last selected formats
    try:
        this_dir = os.path.dirname(os.path.realpath(__file__))
        config_file = os.path.join(this_dir, 'last_export_config.txt')
        with open(config_file, 'w') as f:
            f.write(','.join(selected_formats))
    except:
        ui.messageBox('Failed to save export config.')

    # Read "only visible" checkbox state
    only_visible = inputs.itemById('onlyVisibleBodies').value
    do_export(selected_formats, only_visible)

# Actual export logic
def do_export(selected_formats, only_visible):
    design = app.activeProduct
    exportMgr = design.exportManager
    rootComp = design.rootComponent

    bodies = list(rootComp.bRepBodies)

    if only_visible:
        # Filter only visible bodies
        bodies = [b for b in bodies if b.isVisible]
        
    #ui.messageBox(f'Bodies to export: {len(bodies)}')
    
    if not bodies:
        ui.messageBox('No bodies found to export.')
        return

    # Extract project name and version
    doc_name = app.activeDocument.name
    base_name, _ = os.path.splitext(doc_name)
    version_match = re.search(r'(v\d+)', base_name, re.IGNORECASE)

    if version_match:
        version = version_match.group(1)
        clean_base = base_name[:version_match.start()].rstrip("_ ").strip()
    else:
        version = ''
        clean_base = base_name

    # Clean up names
    safe_base_name = clean_base.replace(" ", "_").replace(".", "_")
    safe_version = version.replace(" ", "_") if version else ''

    # Folder: two levels above script dir → /EXPORTED/<Project>
    this_dir = os.path.dirname(os.path.realpath(__file__))
    parent_dir = os.path.dirname(os.path.dirname(this_dir))
    export_base = os.path.join(parent_dir, 'EXPORTED')
    export_folder = os.path.join(export_base, safe_base_name)

    # Ask user for folder confirmation
    choice = ui.messageBox(
        f'Default export folder:\n{export_folder}\n\nUse this folder?',
        'Export Folder',
        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType
    )

    if choice == adsk.core.DialogResults.DialogNo:
        dlg = ui.createFolderDialog()
        dlg.title = 'Select Export Folder'
        if dlg.showDialog() == adsk.core.DialogResults.DialogOK:
            export_folder = dlg.folder
        else:
            ui.messageBox('Export canceled.')
            return

    if not os.path.exists(export_folder):
        os.makedirs(export_folder)

    # Export each body to each selected format
    for body in bodies:
        safe_body_name = body.name.replace(" ", "_")
        was_visible = body.isVisible  # Remember original visibility
        if not was_visible:
            body.isVisible = True
        
        for fmt in selected_formats:
            filename = (
                f'{safe_body_name}_{safe_version}.{fmt.lower()}'
                if safe_version else f'{safe_body_name}.{fmt.lower()}'
            )
            full_path = os.path.join(export_folder, filename)

            try:
                if fmt == 'STL':
                    try:
                        opts = exportMgr.createSTLExportOptions(body)
                        opts.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementHigh
                        opts.filename = full_path
                        exportMgr.execute(opts)
                    except Exception as e:
                        ui.messageBox(f"Could not create or execute STL export for body '{body.name}':\n{e}")
                        continue
                    opts = exportMgr.createSTLExportOptions(body)
                    opts.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementHigh
                    opts.filename = full_path
                    exportMgr.execute(opts)
                elif fmt == '3MF':
                    opts = exportMgr.createC3MFExportOptions(body)
                    opts.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementHigh
                    opts.filename = full_path
                    exportMgr.execute(opts)
                elif fmt == 'STEP':
                    exportMgr.execute(exportMgr.createSTEPExportOptions(full_path))
                elif fmt == 'IGES':
                    exportMgr.execute(exportMgr.createIGESExportOptions(full_path))
                elif fmt == 'SAT':
                    exportMgr.execute(exportMgr.createSATExportOptions(full_path))
                elif fmt == 'SMT':
                    exportMgr.execute(exportMgr.createSMTExportOptions(full_path))
                elif fmt == 'F3D':
                    exportMgr.execute(exportMgr.createFusionArchiveExportOptions(full_path))
                else:
                    ui.messageBox(f'Unknown format: {fmt}')
            except Exception as e:
                if "InternalValidationError" in str(e):
                    # Optional: silently ignore
                    pass
                else:
                    ui.messageBox(f'Failed to export {fmt} for {safe_body_name}.\n{e}')
        if not was_visible:
            body.isVisible = False  # Restore visibility

    # Ask to open folder
    open_choice = ui.messageBox(
        f'Export completed:\n{export_folder}\n\nOpen folder now?',
        'Export Finished',
        adsk.core.MessageBoxButtonTypes.YesNoButtonType
    )
    if open_choice == adsk.core.DialogResults.DialogYes:
        try:
            if platform.system() == 'Windows':
                os.startfile(export_folder)
            else:
                os.system(f'open "{export_folder}"')
        except Exception as e:
            ui.messageBox(f'Could not open folder:\n{str(e)}')

# Cleanup after command finishes
def command_destroy(args: adsk.core.CommandEventArgs):
    global local_handlers
    local_handlers = []
    futil.log(f'{CMD_NAME} Command Destroy Event')
