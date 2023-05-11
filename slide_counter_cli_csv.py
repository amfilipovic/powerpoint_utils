import csv
import glob
import os
import platform
from pptx import Presentation

# Define a function to clear the console screen based on the OS.
def clear_screen():
    if platform.system() == 'Windows':
        os.system('cls')
    else:
        os.system('clear')

# Define a function to format file size in human-readable format.
def file_format_size(size):
    power = 2**10
    n = 0
    size_labels = {0: 'B', 1: 'KB', 2: 'MB', 3: 'GB', 4: 'TB'}
    while size > power:
        size /= power
        n += 1
    return f'{size:.2f} {size_labels[n]}'

# Get the current folder.
current_folder = os.getcwd()
# Find all PowerPoint presentations in the current folder.
presentations = glob.glob(os.path.join(current_folder, '*.ppt*'))
# Get the script name and create a CSV report with the same name.
script_name = os.path.splitext(os.path.basename(__file__))[0]
report = script_name + '.csv'
# Initialize variables to keep track of the total number of slides and total file size.
total_number_of_slides = 0
total_filesize = 0

# Clear the console screen.
clear_screen()

# Open a CSV file to write the scan results.
with open(report, 'w', newline='', encoding='utf-8-sig') as file:
    # Create a CSV writer object.
    writer = csv.writer(file, delimiter=',')
    # Write the header row to the CSV file.
    writer.writerow([f"Scan results of '{current_folder}':"])
    writer.writerow(['Presentation', 'Number of slides', 'File size'])
    # Loop through each PowerPoint presentation in the current folder.
    for presentation in presentations:
        # Open the presentation using the 'pptx' module.
        current_presentation = Presentation(presentation)
        # Get the number of slides in the presentation.
        number_of_slides = len(current_presentation.slides)
        # Get the file size of the presentation.
        file_size = os.path.getsize(presentation)
        # Write the presentation name, number of slides, and file size to the CSV file.
        writer.writerow([os.path.basename(presentation), number_of_slides, file_format_size(file_size)])
        # Update the total number of slides and total file size variables.
        total_number_of_slides += number_of_slides
        total_filesize += file_size
    # Write the total number of slides and total file size to the CSV file.
    writer.writerow([f"Total number of slides in {len(presentations)} presentation(s): {total_number_of_slides} ({file_format_size(total_filesize)})"])
    # Print the location of the saved report file upon completion.
    print(f"Report saved to '{report}'")