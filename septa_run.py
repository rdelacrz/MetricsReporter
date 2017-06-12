"""
This module runs the script for the GUI that will access Oracle and produce 
metrics.
"""

# Built-in modules
from collections import OrderedDict
from datetime import datetime

# User-defined modules
from constants import *
from calculator import SEPTACalculator
from directories import PROJECT_DIR, RAW_DATA_DIR, PREV_DATA_DIR
from db_accessor.jira_gt import JiraGT
from utilities import create_dirpath, move_old_files
from xl_writer import RawDataWriter, ExportDataWriter

# Constants for data types
MAIN = 'Main Data'
CHANGE = 'Change History'

# Constant for weekday when next date should be omitted
SUNDAY = 6

# Only runs script when it is being directly executed
if (__name__ == '__main__'):
  # Grabs SEPTA data
  generator = JiraGT()
  generator.connect()
  data = generator.get_issue_data('SEPTA')
  generator.disconnect()
  
  # Save path for raw SEPTA data
  raw_save_path = create_dirpath(subdirs=[PROJECT_DIR, 'SEPTA', RAW_DATA_DIR])
    
  # Sets header tuples for raw data sheet
  main_headers = [(KEY, 12), (ISSUETYPE, 16), (PRIORITY, 10), ('Current Status', 16), 
            (CREATED, 17), (RESOLVED, 17), (COMPS, 17), (LINKS, 25), (PACK, 12), 
            (FOUND, 17), (ROOT, 20), (DEV_EST, 14), (DEV_ACT, 14)]
  history_headers = [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 17)]
    
  # Writes data to raw data files
  raw_data = OrderedDict([(MAIN, []), (CHANGE, [])]) 
  for key, params in data.iteritems():
    curr_params = [key]
    for k, v in params.iteritems():
      if (k != HIST): curr_params.append(v)
    raw_data[MAIN].append(curr_params)
    
    # Gets historical data if it exists
    if (HIST in params):
      for index, old in enumerate(params[HIST][OLD]):
        raw_data[CHANGE].append([key, old, params[HIST][NEW][index], params[HIST][TRANS][index]])
  raw_writer = RawDataWriter(raw_data.keys(), [main_headers, history_headers], 
                             'SEPTA Raw Data', raw_save_path)
  raw_writer.produce_workbook(raw_data)
    
  # Perform SEPTA calculations
  calc = SEPTACalculator()
  sev, status, _ = calc.calculate(data, age_data=False)
  
  # Save path for exported SEPTA data
  save_path = create_dirpath(subdirs=[PROJECT_DIR, 'SEPTA'])
  
  # Moves all old Status and Severity files to the old folder
  move_old_files(['Severity', 'Status'], [PROJECT_DIR, 'SEPTA', PREV_DATA_DIR])
  
  # Priority
  series_names = [(p, p) for p in calc.priority_list]
  
  # Populates sheet names map, where each data type corresponds to a workbook
  closure = [TOTAL, OPEN, CLOSED]
  sheet_names_map = OrderedDict([('TOTAL', closure)])
  for data_type in [FAT_A, FAT_B, PILOT, PILOT_HI]:
    sheet_names_map[data_type] = ['%s (%s)' % (x, data_type) for x in closure]
    
  # Iterates through the sheets of each workbook
  for data_type, sheet_names in sheet_names_map.iteritems():
    filename = 'Defects by Severity %s' % data_type
    
    # Backwards because xlsxwriter series colors get reversed based on insertion
    colors = ['blue', 'green', 'yellow', 'orange', 'red']
    fill_map = { series : { 'color' : color } 
              for (series, _), color in zip(series_names, colors) }
    
    # Sets up sheet parameters
    sheet_data = OrderedDict([(sheet, { }) for sheet in sheet_names])
    for sheet in sheet_names:
      sheet_data[sheet]['date_shift'] = 1
      sheet_data[sheet]['omit_last'] = datetime.today().weekday() == SUNDAY
      
      # Specific modifications
      chart_name = '%s %s' % (sheet.split()[0], filename)
      weeks = 26
        
      # Hi Priority pilot should only have three priorities in the series
      if (data_type == PILOT_HI):
        series_names = series_names[0:3]
        colors = ['yellow', 'orange', 'red']
        fill_map = { series : { 'color' : color } 
                for (series, _), color in zip(series_names, colors) }
          
      # Chart parameters
      chart1={'chart_name' : chart_name, 'issue_type':'Defect', 'fill_map' : fill_map}
      chart2={'chart_name' : chart_name, 'issue_type':'Defect', 'fill_map' : fill_map, 
            'insertion_row' : 39 + (2 * len(series_names)), 'weeks' : 26}
      chart3={'chart_name' : chart_name, 'issue_type':'Defect', 'fill_map' : fill_map, 
            'insertion_row' : 70 + (2 * len(series_names)), 'weeks' : 9}
      
      # Sets chart parameters
      sheet_data[sheet]['chart_params'] = [chart1, chart2, chart3]
      
    # Performs exports
    exporter = ExportDataWriter(filename, save_path, series_names)
    exporter.produce_workbook(sev, sheet_data=sheet_data)
  
  # Status
  series_names = calc.get_status_desc()
  sheet_names = [TOTAL, FAT_1A, FAT_1B, FAT_1B_HI, PILOT, PILOT_HI]
  filename = 'Defects by Status'
      
  # Set up sheet parameters
  sheet_data = OrderedDict([(sheet, { }) for sheet in sheet_names])
  for sheet in sheet_names:
    sheet_data[sheet]['date_shift'] = 1
    sheet_data[sheet]['omit_last'] = datetime.today().weekday() == SUNDAY
    
    # Specific modifications
    chart_name = filename
    weeks = 26
    min_y = None
    if (sheet == TOTAL): min_y = 3100
    elif (sheet == PILOT): chart_name = 'All Defects by Status (Pilot)'
    elif (sheet == PILOT_HI):
      chart_name = 'Hi Priority Defects by Status (Pilot)'
      weeks = 9
      min_y = 600
    
    # Chart parameters
    chart1={ 'chart_name' : chart_name, 'issue_type' : 'Defect' }
    chart2={'chart_name' : chart_name, 'issue_type' : 'Defect', 'min_y' : min_y,
          'insertion_row' : 39 + (2 * len(series_names)), 'weeks' : weeks}
    
    # Sets chart parameters
    sheet_data[sheet]['chart_params'] = [chart1, chart2]
      
  # Performs export
  exporter = ExportDataWriter(filename, save_path, series_names)
  exporter.produce_workbook(status, sheet_data=sheet_data)