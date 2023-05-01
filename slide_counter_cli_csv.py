import csv
import glob
import os
import platform
from pptx import Presentation

def clear_screen():
    if platform.system() == 'Windows':
        os.system('cls')
    else:
        os.system('clear')

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

clear_screen()

with open(report, 'w', newline='', encoding='utf-8-sig') as file:
    writer = csv.writer(file, delimiter=',')
    writer.writerow([f"Scan results of '{current_folder}':"])
    writer.writerow(['Presentation', 'Number of slides', 'File size'])
    for presentation in presentations:
        current_presentation = Presentation(presentation)
        number_of_slides = len(current_presentation.slides)
        file_size = os.path.getsize(presentation)
        writer.writerow([os.path.basename(presentation), number_of_slides, file_format_size(file_size)])
        total_number_of_slides += number_of_slides
        total_filesize += file_size
    writer.writerow([f"Total number of slides in {len(presentations)} presentation(s): {total_number_of_slides} ({file_format_size(total_filesize)})"])
    print(f"Report saved to '{report}'")