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

current_folder = os.getcwd()
presentations = glob.glob(os.path.join(current_folder, '*.ppt*'))
script_name = os.path.splitext(os.path.basename(__file__))[0]
report = script_name + '.html'
total_number_of_slides = 0
total_filesize = 0

with open(report, 'w', encoding='utf-8') as file:
    file.write('<!DOCTYPE html>\n')
    file.write('<html>\n')
    file.write('<head>\n')
    file.write('<meta charset="UTF-8">\n')
    file.write('<title>Slide Counter</title>\n')
    file.write('</head>\n')
    file.write('<body>\n')
    header = f"Scan results of '{current_folder}':"
    file.write('<h3>{}</h3>\n'.format(header))
    file.write('<table>\n')
    file.write('<tr><th>Presentation</th><th>Number of slides</th><th>File size</th></tr>\n')
    for presentation in presentations:
        file_size = os.path.getsize(presentation)
        current_presentation = Presentation(presentation)
        number_of_slides = len(current_presentation.slides)
        file.write('<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n'.format(os.path.basename(presentation), number_of_slides, file_format_size(file_size)))
        total_number_of_slides += number_of_slides
        total_filesize += file_size
    file.write('</table>\n')
    footer = f"Total number of slides in {len(presentations)} presentation(s): {total_number_of_slides} ({file_format_size(total_filesize)})"
    file.write('<h4>{}</h4>\n'.format(footer))
    file.write('</body>\n')
    file.write('</html>')
    print(f"Report saved to '{report}'")