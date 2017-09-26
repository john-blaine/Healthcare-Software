#!/usr/bin/env python

'''
Postcondition: Create a program to read a CSV file, take input from user,
and from this create a view of the data which will then be saved to a CSV file.

Subgoal 1: Read data from CSV file into memory and store it in a variable

Subgoal 2: Get input from user for filtering the data read from CSV file

Subgoal 3: Filter data using the inputted user data

Subgoal 4: Get input from user for sorting the filtered data

Subgoal 5: Sort the filtered data

Subgoal 6: Get input from user for naming file to save filtered data

Subgoal 7: Save filtered data to file using inputted user data

'''

import csv
import os
import sys

class FileHeadersNotReconciled(Exception): pass
class ValueNotInteger(Exception): pass

csv_column_names = ('Zip Code',
                    'Total Population',
                    'Median Age',
                    'Total Males',
                    'Total Females',
                    'Total Households',
                    'Average Household Size'
                    )
# Subgoal 1
def load_census_data(filepath):
    csv_data_list = []

    with open(file_path, 'r') as csvfile:
        for row in csv.DictReader(csvfile):
            for k, v in row.items():
                valid_data = True
                try:
                    if k not in csv_column_names:
                        raise FileHeadersNotReconciled
                except FileHeadersNotReconciled:
                    print('Column names in row missing or invalid.')
                    valid_data = False
                    break
                try:
                    if not v.isdigit():
                        raise ValueNotInteger
                except ValueNotInteger:
                    print('One or more values in row are not integers' \
                        + ' and cannot be converted into integers')
                    valid_data = False
                    break
            if valid_data:
                csv_data_list.append(row)
                
        csv_data = tuple(csv_data_list)
        return csv_data
    
# Subgoal 3
def filter_data_by_column_and_floor(csv_data, column_name, floor_value):
    new_csv_data_list = []
    for row in csv_data:
        valid_data = True
        for k, v in row.items():
            if k == column_name and int(v) < floor_value:
                valid_data = False
        if valid_data:
            for k, v in row.items():
                row[k] = int(v)
            new_csv_data_list.append(row)
    new_csv_data = tuple(new_csv_data_list)
    return new_csv_data

file_path = os.path.join(os.path.realpath('data'),
                         ('2010_Census_Populations_by_Zip_Code.csv'))

census_data = load_census_data(file_path)

# Subgoal 2
for column_number, column_name in enumerate(csv_column_names):
    print('{}) {}'.format(column_number + 1, column_name))
print('')

attempts = 0

while (attempts < 3):
    try:
        user_input_column = int(input('Please enter the number of the '
                               + 'column you wish to filter the data by: '))
        print('')
    except ValueError:
        print('\nInvalid input. Please enter an integer corresponding to one '
              + 'of the listed columns (1-7). \n')
        attempts += 1
        continue
    if user_input_column > 7 or user_input_column < 1:
        print('That input is invalid. Please enter an integer between 1 and 7. \n')
        attempts += 1
        continue
    break
else:
    print('Too many attempts were made. Exiting.')
    sys.exit()

attempts = 0

while (attempts < 3):
    try:
        user_input_floor = int(input('Please enter the floor value you wish the data '
                               + 'to be filtered by: '))
        print('')
    except ValueError:
        print('\nInvalid input. Please enter an integer.\n')
        attempts += 1
        continue
    break
else:
    print('Too many attempts were made. Exiting.')
    sys.exit()

new_census_data = filter_data_by_column_and_floor(census_data,
                                                  csv_column_names[user_input_column-1],
                                                  user_input_floor
                                                  )
# Subgoal 4

attempts = 0

while (attempts < 3):
    try:
        user_input_sort_column = int(input('Please enter the number of the '
                               + 'column you wish to sort the data by: '))
        print('')
    except ValueError:
        print('\nInvalid input. Please enter an integer corresponding to one '
              + 'of the listed columns (1-7). \n')
        attempts += 1
        continue
    if user_input_sort_column > 7 or user_input_sort_column < 1:
        print('That input is invalid. Please enter an integer between 1 and 7. \n')
        attempts += 1
        continue
    break
else:
    print('Too many attempts were made. Exiting.')
    sys.exit()

# Subgoal 5

new_census_data = sorted(new_census_data, key = lambda d: d[csv_column_names[user_input_sort_column-1]])

# Subgoal 6

attempts = 0

while (attempts < 3):
    try:
        final_file_name = input('Please enter the name you '
                               + 'wish to save the file as '
                               + '(please do not include .csv extension): ' )
        print('')
    except ValueError:
        print('Invalid input. Please enter a valid string.\n')
        attempts += 1
        continue
    if not user_input_sort_column:
        print('You have entered an empty string. Please enter a valid string.\n')
        attempts += 1
        continue
    if os.path.isfile(os.path.join(os.path.realpath('exports'),
                         (final_file_name +'.csv'))):
        print('The filename you have entered already exists. Please try again.\n')
        attempts += 1
        continue
    break
else:
    print('Too many attempts were made. Exiting.')
    sys.exit()

# Subgoal 7

new_file_path = os.path.join(os.path.realpath('exports'),
                         (final_file_name +'.csv'))

with open(new_file_path, 'w') as newcsvfile:

    write_new_csv_file = csv.DictWriter(newcsvfile, fieldnames=(csv_column_names),
                                        lineterminator = '\n')
    write_new_csv_file.writeheader()
    for row in new_census_data:
        write_new_csv_file.writerow(row)

