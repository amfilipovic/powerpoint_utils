import os
from comtypes import client
from pptx import Presentation

# Set the path to the current folder.
path = os.getcwd()

# Create PowerPoint application object.
powerpoint_app = client.CreateObject("Powerpoint.Application")

# Loop through presentations in current folder.
for filename in os.listdir(path):
    if filename.endswith(".pptx") or filename.endswith(".ppt"):
        # Set the full path of the presentation.
        filepath = os.path.join(path, filename)
        # Open the presentation.
        presentation = powerpoint_app.Presentations.Open(filepath)
        # Set the name of the PDF output file.
        outputpdfname = os.path.splitext(filename)[0] + ".pdf"
        # Set the path of the PDF output file.
        outputpdfpath = os.path.join(path, outputpdfname)
        # Create the PDF output file.
        presentation.ExportAsFixedFormat(outputpdfpath, 2, PrintRange=None)
        # Close the presentation.
        presentation.Close()

# Close the PowerPoint application.
powerpoint_app.Quit()