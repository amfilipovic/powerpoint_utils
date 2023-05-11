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

current_folder = os.getcwd()
presentations = glob.glob(os.path.join(current_folder, '*.ppt*'))
script_name = os.path.splitext(os.path.basename(__file__))[0]
report = script_name + '.csv'
total_number_of_slides = 0
total_filesize = 0

# Clear the console screen.
clear_screen()

current_folder = os.getcwd()
presentations = glob.glob(os.path.join(current_folder, '*.ppt*'))
script_name = os.path.splitext(os.path.basename(__file__))[0]
report = script_name + '.txt'
total_number_of_slides = 0
total_filesize = 0

with open(report, 'w') as file:
    file.write(f"Scan results of '{current_folder}':\n")
    for presentation in presentations:
        current_presentation = Presentation(presentation)
        number_of_slides = len(current_presentation.slides)
        file_size = os.path.getsize(presentation)
        file.write(f"Number of slides in '{os.path.basename(presentation)}': {number_of_slides} ({file_format_size(file_size)})\n")
        total_number_of_slides += number_of_slides
        total_filesize += file_size
    file.write(f"Total number of slides in {len(presentations)} presentation(s): {total_number_of_slides} ({file_format_size(total_filesize)})")
    print(f"Report saved to '{report}'")