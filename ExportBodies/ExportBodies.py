# Author-Patrick Rainsberry
# Description-A sample Fusion Addin to demonstrate various UI elements.

# Assuming you have not changed the general structure of the template no modification is needed in this file.
from . import commands
from .lib import fusionAddInUtils as futil
import adsk.core


def run(context):
    try:
        # Display a message when the add-in is manually run.
        if not context['IsApplicationStartup']:
            app = adsk.core.Application.get()
            ui = app.userInterface
            ui.messageBox(
                'This tool exports all visible bodies in the root component to selected file formats.\n\n'
                '✔ Select one or more export formats (e.g., STL, STEP, IGES).\n'
                '✔ You will be asked to confirm or change the export folder.\n'
                '✔ All exported files will include version info (if present in filename).\n\n'
                'After exporting, you can choose to open the folder automatically.',
                '🧰 Export Bodies – Instructions'
            )
    
        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.start()

    except:
        futil.handle_error('run')


def stop(context):
    try:
        # Remove all of the event handlers your app has created
        futil.clear_handlers()

        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.stop()

    except:
        futil.handle_error('stop')