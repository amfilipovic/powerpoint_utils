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

# Clear the console screen.
clear_screen()

current_folder = os.path.dirname(os.path.abspath(__file__))
presentations = glob.glob(os.path.join(current_folder, '*.ppt*'))
total_number_of_slides = 0
total_filesize = 0

print(f"Scan results of '{current_folder}':")
for presentation in presentations:
    current_presentation = Presentation(presentation)
    number_of_slides = len(current_presentation.slides)
    file_size = os.path.getsize(presentation)
    print(f"Number of slides in '{os.path.basename(presentation)}': {number_of_slides} ({file_format_size(file_size)})")
    total_number_of_slides += number_of_slides
    total_filesize += file_size
if presentations:
    print(f"Total number of slides in {len(presentations)} presentation(s): {total_number_of_slides} ({file_format_size(total_filesize)})")
else:
    clear_screen()
    print("Error: There are no presentations in this folder.")