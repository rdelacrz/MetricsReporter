"""
This module runs the script for the GUI that will access Oracle and produce 
metrics.
"""

# Built-in modules
from collections import OrderedDict

# User-defined modules
from constants import *
from directories import PROJECT_DIR, RAW_DATA_DIR
from db_accessor.jira_gt import JiraGT
from utilities import create_dirpath
from xl_writer import RawDataWriter

# Constants for data types
MAIN = 'Main Data'
CHANGE = 'Change History'

# Only runs script when it is being directly executed
def create_raw_reports(project_list):
  project_data = OrderedDict()
  
  # Connects to JIRA
  generator = JiraGT()
  generator.connect()
  
  # Iterates through each project and queries for its data
  for proj in project_list:
    project_data[proj] = generator.get_roche_issue_data(proj, 
                          ['Defect','Defect Subtask','Change Request','CR Sub Task','Task','Task Sub Task', 
                           'New Development', 'Compliance', 'Proposed Work', 'Action Item', 'Incident'])

    # Save path for raw SEPTA data
    raw_save_path = create_dirpath(subdirs=[PROJECT_DIR, proj, RAW_DATA_DIR])
      
    # Sets header tuples for raw data sheet
    main_headers = [('Row Number', 3), (KEY, 12), (ISSUETYPE, 16), (PRIORITY, 10), ('Current Status', 16), 
              (CREATED, 17), (RESOLVED, 17), (COMPS, 17), (LINKS, 25), (PACK, 12), 
              (FOUND, 17), (PBI, 20), (ROOT, 20), (DEV_EST, 14), (DEV_ACT, 14)]
    history_headers = [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 17)]
      
    # Writes data to raw data files
    raw_data = OrderedDict([(MAIN, []), (CHANGE, [])]) 
    for key, params in project_data[proj].iteritems():
      curr_params = [key]
      for k, v in params.iteritems():
        if (k != HIST): curr_params.append(v)
      raw_data[MAIN].append(curr_params)
      
      # Gets historical data if it exists
      if (HIST in params):
        for index, old in enumerate(params[HIST][OLD]):
          raw_data[CHANGE].append([key, old, params[HIST][NEW][index], params[HIST][TRANS][index]])
    raw_writer = RawDataWriter(raw_data.keys(), [main_headers, history_headers], 
                               '%s Raw Data' % proj, raw_save_path)
    raw_writer.produce_workbook(raw_data)
    
    print "%s raw data produced" % proj

  # Disconnects from JIRA
  generator.disconnect()

# Only runs script when it is being directly executed
if (__name__ == '__main__'):
  TOLL = ['BAIFA', 'BATAVEC', 'FLCSS', 'FTB', 'MDCTC', 'MDEOF', 'NCTABOS', 'NCTARTCS', 'NHDOT', 'NJND', 'NYAET', 'TXDOT']
  project_list = ['SUN']
  create_raw_reports(project_list)