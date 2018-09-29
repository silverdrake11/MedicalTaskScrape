import csv
import os.path
import re
import traceback
from datetime import datetime

from html2text import html2text


# Be careful modifying this as these are the exact names of the tag fields that are pulled out from the table
FIELDS = ['NumberPhysicians', 'Attending', 'APRN', 'NPPN', 'PGY', 'Hospitalist', 'Interaction', 'Comments']


# Do not modify
CHUNK_START = 'Task:'
CHUNK_END = '---|---'


def add_team_size(merged_data):
    '''Adds team size which sums some numeric fields to determine the number of people present in the team'''

    total = 0
    for item in ['NumberPhysicians', 'APRN', 'NPPN', 'PGY', 'Hospitalist']:
        if item in merged_data:
            value = merged_data[item]
            try:
                value = float(value)
            except ValueError:
                value = 0
            total += value
    if total:
        merged_data['TeamSize'] = total



def clean_up_some_fields(fields):
    '''The field data is generally a bit rougher to work with so additional processing needs to be done on some of the fields'''

    for field in fields:

        # Remove the 'PAs'
        if field == 'APRN':
            aprn = fields[field]
            fields[field] = re.sub(r"\s*PAs?", "", aprn)

        # Ignore if not a number
        if field == 'Interaction' or field == 'Hospitalist' or field == 'PGY' or field == 'NPPN':
            numeric_field = fields[field]
            try:
                float(numeric_field)
            except ValueError:
                fields[field] = ''

        # Strip newlines as this causes problems with CSV files
        if field == 'Comments':
            comments = fields[field]
            comments.replace('\n', ' ').replace('\r', '')
            fields[field] = comments


def get_field_data(text):
    '''
    This is the part of the table where the information is generally enclosed in tags </>
    which are defined in the global variable FIELDS
    '''

    fields = {}
    for field in FIELDS:

        pattern = r'<?/?\(?/?' + field + r'>?\)?' # Some of the tags are with parenthesis and some are missing random parts (typos) This regex catches most of them (not all)
        split_text = re.split(pattern, text)

        # Get the contents inside a tag and strip the whitespace
        if len(split_text) > 2:
            contents = split_text[1].strip()
            if contents:
                fields[field] = contents

    return fields


def get_task_data(task_parts):
    '''This is the first part of the table which has information such as Name, Date and Hours'''

    # Remove whitespace and newlines
    cleaned = []
    for part in task_parts:
        part = part.strip()
        if part:
            cleaned.append(part)

    task1 = cleaned[0].replace('Task:', '').strip()
    tracking, task2 = cleaned[1].split('>')
    tracking = tracking.strip()
    tracking = tracking.split(' ')[0]

    hours = cleaned[-2]

    _, name, date = cleaned[-1].split('_')

    name = name.split('(')[0].strip() # The name includes parenthesis and the username so get rid of that

    # Read the date into a Python datetime
    date = date.strip()[:-4]
    try:
        date_obj = datetime.strptime(date, '%b %d, %Y %I:%M:%S %p')
    except Exception:
        print(traceback.format_exc())
        print("WARNING: Unable to parse '{}'\n".format(date))
        date_obj = None

    task_data = {}

    if name:
        task_data['Name'] = name
    if date_obj:
        task_data['Date'] = date_obj.strftime("%m/%d/%Y") # Write out the datetime object
    if hours:
        task_data['Hours'] = hours
    if tracking:
        task_data['Tracking'] = tracking
    if task1:
        task_data['Dept'] = task1

    return task_data


def get_data(chunk):
    '''Given a chunk returns a dictionary with relevant information if available'''

    task_parts = chunk.split('|')
    last_part = task_parts.pop().strip() # Careful this modifies task_parts

    tasks = get_task_data(task_parts)

    fields = get_field_data(last_part)
    clean_up_some_fields(fields)

    merged_data = dict(list(tasks.items()) + list(fields.items()))
    add_team_size(merged_data)

    return merged_data


def get_chunks(lines):
    '''
    Split the text data from the html2text into chunks. Chunks start with 'Task:' and end
    with ---|--- which designates the end of a table in html2text
    '''

    start = False
    chunks = []
    chunk = None

    for cur_line in lines:

        cur_line = cur_line.strip()

        if re.match(r'^\s*\|\s+'+CHUNK_START, cur_line):
            start = True
            chunk = []

        if cur_line == CHUNK_END:
            start = False

            if chunk:
                chunk_str = ' '.join(chunk)
                chunks.append(chunk_str)

        if start:
            chunk.append(cur_line)

    return chunks


def write_csv(input_filepath, output_filepath):

    # Get text from HTML using html2text
    with open(input_filepath) as fh:
        html = fh.read()
        html_text = html2text(html)

    chunks = get_chunks(html_text.splitlines())

    with open(output_filepath, 'w', newline='') as csvfile:

        first_part = ['Name', 'Date', 'Hours', 'Tracking', 'Dept', 'TeamSize']
        fieldnames = first_part + FIELDS

        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for chunk in chunks:

            if 'Attending' not in chunk: # This does a pretty good job at filtering out irrelevant information
                continue

            csv_row = {}
            
            chunk_data = get_data(chunk)

            # Add items to the csv row (for DictWriter) if they exist
            for item in fieldnames:
                if item in chunk_data:
                    cell = chunk_data[item]
                    if isinstance(cell, str): # Semicolons and tabs seem to cause a problem delimiting the CSV
                        cell = cell.replace(';',',').replace(r'\t','') 
                    csv_row[item] = cell

            writer.writerow(csv_row)







