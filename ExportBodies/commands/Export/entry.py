import adsk.core
import adsk.fusion
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
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    inputs = args.command.commandInputs
    group = inputs.addGroupCommandInput('formatGroup', 'Export Formats')
    format_inputs = group.children

    # Try loading previous export format config
    this_dir = os.path.dirname(os.path.realpath(__file__))
    config_file = os.path.join(this_dir, 'last_export_config.txt')
    last_formats = []
    last_refinement = 'High'
    last_surface_dev  = 0.0005
    last_normal_dev   = 0.17453
    last_max_edge     = 0.2
    last_aspect_ratio = 21.5
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            lines = f.read().splitlines()
            last_formats = [x.strip().upper() for x in lines[0].split(',')]
            if len(lines) > 1 and lines[1].strip() in ('Low', 'Medium', 'High', 'Custom'):
                last_refinement = lines[1].strip()
            if len(lines) > 2:
                try:
                    vals = [float(x) for x in lines[2].split(',')]
                    if len(vals) == 4:
                        last_surface_dev, last_normal_dev, last_max_edge, last_aspect_ratio = vals
                except:
                    pass
    else:
        last_formats = ['STL', '3MF']

    # Add format checkboxes
    formats = ['STL', '3MF', 'STEP', 'IGES', 'SAT', 'SMT']
    for fmt in formats:
        format_inputs.addBoolValueInput(fmt, fmt, True, '', fmt in last_formats)

    # Collapsible group: Body Selection
    body_group = inputs.addGroupCommandInput('bodyGroup', 'Body Selection')
    body_group.isExpanded = True
    bg = body_group.children

    bg.addBoolValueInput('onlyVisibleBodies', 'Only export visible bodies', True, '', True)
    bg.addBoolValueInput('exportComponents', 'Export bodies from components', True, '', False)

    # Collapsible group: Mesh Refinement (STL & 3MF)
    mesh_group = inputs.addGroupCommandInput('meshGroup', 'Mesh Refinement (STL & 3MF)')
    mesh_group.isExpanded = True
    mg = mesh_group.children

    refinement_input = mg.addDropDownCommandInput('meshRefinement', 'Refinement', adsk.core.DropDownStyles.LabeledIconDropDownStyle)
    refinement_items = refinement_input.listItems
    refinement_items.add('Low',    last_refinement == 'Low',    '')
    refinement_items.add('Medium', last_refinement == 'Medium', '')
    refinement_items.add('High',   last_refinement == 'High',   '')
    refinement_items.add('Custom', last_refinement == 'Custom', '')

    is_custom = last_refinement == 'Custom'

    surf_input = mg.addFloatSpinnerCommandInput('surfaceDeviation', 'Surface Deviation (cm)', 'cm', 0.00001, 10.0, 0.0001, last_surface_dev)
    surf_input.isVisible = is_custom

    norm_input = mg.addFloatSpinnerCommandInput('normalDeviation', 'Normal Deviation (rad)', '', 0.001, 3.14159, 0.01, last_normal_dev)
    norm_input.isVisible = is_custom

    edge_input = mg.addFloatSpinnerCommandInput('maximumEdgeLength', 'Max Edge Length (cm)', 'cm', 0.0001, 100.0, 0.01, last_max_edge)
    edge_input.isVisible = is_custom

    aspect_input = mg.addFloatSpinnerCommandInput('aspectRatio', 'Aspect Ratio', '', 1.0, 1000.0, 0.5, last_aspect_ratio)
    aspect_input.isVisible = is_custom

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

    # Read mesh refinement settings
    selected_refinement_name = inputs.itemById('meshRefinement').selectedItem.name
    surface_dev   = inputs.itemById('surfaceDeviation').value
    normal_dev    = inputs.itemById('normalDeviation').value
    max_edge      = inputs.itemById('maximumEdgeLength').value
    aspect_ratio  = inputs.itemById('aspectRatio').value

    # Save last selected formats, refinement and custom values
    try:
        this_dir = os.path.dirname(os.path.realpath(__file__))
        config_file = os.path.join(this_dir, 'last_export_config.txt')
        with open(config_file, 'w') as f:
            f.write(','.join(selected_formats) + '\n')
            f.write(selected_refinement_name + '\n')
            f.write(f'{surface_dev},{normal_dev},{max_edge},{aspect_ratio}')
    except:
        ui.messageBox('Failed to save export config.')

    # Read checkbox states
    only_visible = inputs.itemById('onlyVisibleBodies').value
    export_components = inputs.itemById('exportComponents').value

    refinement_map = {
        'Low':    adsk.fusion.MeshRefinementSettings.MeshRefinementLow,
        'Medium': adsk.fusion.MeshRefinementSettings.MeshRefinementMedium,
        'High':   adsk.fusion.MeshRefinementSettings.MeshRefinementHigh,
        'Custom': adsk.fusion.MeshRefinementSettings.MeshRefinementCustom,
    }
    mesh_refinement = refinement_map.get(selected_refinement_name, adsk.fusion.MeshRefinementSettings.MeshRefinementHigh)

    do_export(selected_formats, only_visible, export_components, mesh_refinement,
              surface_dev, normal_dev, max_edge, aspect_ratio)

# Called when any input changes – show/hide custom sliders
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed = args.input
    inputs = args.inputs
    if changed.id == 'meshRefinement':
        is_custom = changed.selectedItem.name == 'Custom'
        inputs.itemById('surfaceDeviation').isVisible  = is_custom
        inputs.itemById('normalDeviation').isVisible   = is_custom
        inputs.itemById('maximumEdgeLength').isVisible = is_custom
        inputs.itemById('aspectRatio').isVisible       = is_custom

# Actual export logic
def do_export(selected_formats, only_visible, export_components=False,
             mesh_refinement=adsk.fusion.MeshRefinementSettings.MeshRefinementHigh,
             surface_dev=None, normal_dev=None, max_edge=None, aspect_ratio=None):
    design = app.activeProduct
    exportMgr = design.exportManager
    rootComp = design.rootComponent

    # Collect all bodies with metadata
    all_bodies = []
    
    # 1. Root component bodies
    for body in rootComp.bRepBodies:
        if not only_visible or body.isVisible:
            all_bodies.append({
                'body': body,
                'name': body.name,
                'component': 'Root'
            })
    
    # 2. Component bodies (if enabled)
    if export_components:
        for occ in rootComp.allOccurrences:
            comp = occ.component
            for body in comp.bRepBodies:
                if not only_visible or body.isVisible:
                    # Component name + body name for unique naming
                    comp_name = comp.name.replace(" ", "_").replace(".", "_")
                    body_name = body.name.replace(" ", "_").replace(".", "_")
                    full_name = f"{comp_name}_{body_name}"
                    
                    all_bodies.append({
                        'body': body,
                        'name': full_name,
                        'component': comp.name
                    })

    if not all_bodies:
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
    for body_info in all_bodies:
        body = body_info['body']
        safe_body_name = body_info['name'].replace(" ", "_")
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
                    opts = exportMgr.createSTLExportOptions(body)
                    opts.meshRefinement = mesh_refinement
                    if mesh_refinement == adsk.fusion.MeshRefinementSettings.MeshRefinementCustom:
                        if surface_dev  is not None: opts.surfaceDeviation  = surface_dev
                        if normal_dev   is not None: opts.normalDeviation   = normal_dev
                        if max_edge     is not None: opts.maximumEdgeLength = max_edge
                        if aspect_ratio is not None: opts.aspectRatio       = aspect_ratio
                    opts.filename = full_path
                    exportMgr.execute(opts)
                elif fmt == '3MF':
                    opts = exportMgr.createC3MFExportOptions(body)
                    opts.meshRefinement = mesh_refinement
                    if mesh_refinement == adsk.fusion.MeshRefinementSettings.MeshRefinementCustom:
                        if surface_dev  is not None: opts.surfaceDeviation  = surface_dev
                        if normal_dev   is not None: opts.normalDeviation   = normal_dev
                        if max_edge     is not None: opts.maximumEdgeLength = max_edge
                        if aspect_ratio is not None: opts.aspectRatio       = aspect_ratio
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
