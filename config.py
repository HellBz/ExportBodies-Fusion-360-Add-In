# Application Global Variables
# Adding application wide global variables here is a convenient technique
# It allows for access across multiple event handlers and modules

# Set to False to remove most log messages from text palette
import os

# ******** App Configuration Options ********
DEBUG = True

ADDIN_NAME = os.path.basename(os.path.dirname(__file__))

COMPANY_NAME = 'HellBz'

# ******** General Configuration of App UI and button placement ********
design_workspace = 'FusionSolidEnvironment'
tools_tab_id = "SolidTab"
tools_tab_name = 'SOLID'

my_tab_name = "Export"  # Only used if creating a custom Tab

my_panel_id = f'{ADDIN_NAME}_panel_2'
my_panel_name = 'Export Bodies'
my_panel_after = ''
