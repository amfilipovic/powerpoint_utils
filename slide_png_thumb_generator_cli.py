import os
import platform
import win32com.client

# Define a function to clear the console screen.
def clear_screen():
    if platform.system() == 'Windows':
        os.system('cls')
    else:
        os.system('clear')

# Call the clear_screen() function to clear the console screen.
clear_screen()

# Create an instance of the PowerPoint application.
powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")

# Get the current working directory.
current_folder = os.getcwd()

# Set the source folder to the current working directory.
source_folder = current_folder

# Iterate over all files in the source folder.
for filename in os.listdir(source_folder):
    # Check if the file is a PowerPoint presentation (.pptx or .ppt).
    if filename.endswith(".pptx") or filename.endswith(".ppt"):
        # Get the full path of the file.
        full_path = os.path.join(source_folder, filename)
        # Open the presentation in PowerPoint.
        presentation = powerpoint_app.Presentations.Open(full_path)
        # Get the first slide of the presentation.
        slide = presentation.Slides(1)
        # Calculate the width and height of the thumbnail based on the aspect ratio of the slide's master slide.
        width_pts = slide.Master.Width
        height_pts = slide.Master.Height
        thumbnail_width = int(960 * (width_pts / height_pts))
        thumbnail_height = 960
        # Export the slide as a PNG image with the filename consisting of the original filename with the "_thumb.png" suffix added.
        slide.Export(os.path.join(source_folder, f"{filename}_thumb.png"), "png", thumbnail_width, thumbnail_height)
        # Close the presentation.
        presentation.Close()

# Quit the PowerPoint application.
powerpoint_app.Quit()