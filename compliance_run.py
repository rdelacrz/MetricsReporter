"""
This module runs the script for the GUI that will access Oracle and produce 
metrics.
"""

# Built-in modules
from collections import OrderedDict
from time import time

# User-defined modules
from constants import *
from calculator import ComplianceCalculator
from directories import PROJECT_DIR, RAW_DATA_DIR, PREV_DATA_DIR
from db_accessor.jira_gt import JiraGT
from utilities import create_dirpath, move_old_files
from xl_writer import RawDataWriter, ExportDataWriter

def timer(func):
  """
  A decorator function that prints the execution time of whatever function is
  passed into it.
  """
  
  def func_wrapper(action='Operation', *args, **kwargs):
    """
    Function wrapper that prints the execution time of the func passed by
    timer() after executing func.
    
    @param action: Textual description of the action being performed by the function.
    @param *args: Arbitrary parameters passed into function.
    @param **kwargs: Arbitrary keyword parameters passed into function.
    """
    
    start_time = time()
    results = func(*args, **kwargs)
    print "\n>>>> %s ran in %f seconds. <<<<\n" % (action, time() - start_time)
    return results
  
  return func_wrapper

@timer
def query_data():
  """
  Queries Compliance data from the JIRA back end database and returns its results.
  
  @return: Raw data queried from the JIRA back end database.
  """
  
  # Access JIRA backend and gets data
  db_accessor = JiraGT()
  db_accessor.connect()
  data = db_accessor.get_compliance_data()
  db_accessor.disconnect()
  return data

@timer
def generate_raw_data(data):
  """
  Re-arranges the data pulled from the back end of the JIRA database and 
  generates raw data Excel workbooks containing it.
  
  @param data: Data pulled from the back end of the JIRA database. It has the
  following structure:
    <key> -> <data type> -> <data value>
          -> <HIST> -> <OLD, NEW, TRANS> -> <data value>
  """
  
  # Gets raw data path
  raw_save_path = create_dirpath(subdirs=[PROJECT_DIR, 'Compliance', RAW_DATA_DIR])
    
  # Sets header tuples for raw data sheet
  main_headers = [(KEY, 12), (PROJECT, 16), (ISSUETYPE, 16), (PRIORITY, 10), ('Current Status', 16), 
            (CREATED, 17), (RESOLVED, 17), (COMPS, 17), (LINKS, 25), (PACK, 12), 
            (DEV_EST, 14), (DEV_ACT, 14)]
  history_headers = [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 25)]
    
  # Configures data dictionary for writing to the Excel workbooks
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
        
  # Writes configured raw data to Excel workbook
  raw_writer = RawDataWriter(raw_data.keys(), [main_headers, history_headers], 
                             'Compliance Raw Data', raw_save_path)
  raw_writer.produce_workbook(raw_data)
  
@timer
def do_calculations(data):
  """
  Performs calculations on raw data pulled from the JIRA back end, and returns
  two data dictionaries containing the resulting calculations. 
  
  @param data: Data pulled from the back end of the JIRA database. It has the
  following structure:
    <key> -> <data type> -> <data value>
          -> <HIST> -> <OLD, NEW, TRANS> -> <data value>
  @return: A reference to the calculator object itself, followed by the data 
  dictionaries containing the resulting calculations. Each data dictionary has
  the following structure:
    <date> -> <top level category> -> <data type> -> <count>
  """
  
  # Perform Compliance calculations
  calc = ComplianceCalculator()
  sev, status, _ = calc.get_metrics(data)
  return calc, sev, status

@timer
def export_metrics(calc, sev, status):
  """
  Takes the results of the calculated metrics and exports them to Excel
  workbooks.
  
  @param calc: Calculator responsible for determining the severity and status
  data.
  @param sev: Data dictionary containing the calculated severity data.
  @param status: Data dictionary containing the calculated status data.
  """
  
  # Save path for exported Compliance data
  save_path = create_dirpath(subdirs=[PROJECT_DIR, calc.project])
  
  # Moves all old Status and Severity files to the old folder
  move_old_files(['Severity', 'Status'], [PROJECT_DIR, calc.project, 
                                          PREV_DATA_DIR])
  # Performs severity and status exports
  projects = status[status.keys()[0]].keys()
  for data, series_names, sheet_names, filename in [
      (sev, [(p, p) for p in calc.priority_list], [TOTAL, OPEN, CLOSED], 
       'Compliance by Severity'),
      (status, calc.get_status_desc(), projects, 'Compliance by Status')
  ]:
    # Sets sheet parameters
    sheet_data = {
      sheet : { 
        'omit_last' : False, 'chart_params' : [
          {'chart_name' : filename, 'issue_type':'Compliance'}, 
          {'chart_name' : filename, 'issue_type':'Compliance', 
           'insertion_row' : 39 + (2 * len(series_names)), 'weeks' : 26 }
        ]
      } for sheet in sheet_names
    }
    
    # Performs export
    exporter = ExportDataWriter(filename, save_path, series_names)
    exporter.produce_workbook(data, sheet_data=sheet_data)

# Only runs script when it is being directly executed
if (__name__ == '__main__'):
  # Grabs Compliance data
  data = query_data(action='Querying Compliance data')
  
  # Generates raw Compliance data workbooks
  generate_raw_data(data=data, action='Producing raw data')
    
  # Calculates severity and status data
  calc, sev, status = do_calculations(data=data, action='Performing calculations')
  
  # Exports severity and status data
  export_metrics(calc=calc, sev=sev, status=status, action='Exporting metric data')