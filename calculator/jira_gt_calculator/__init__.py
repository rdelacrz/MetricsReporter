"""
This module contains the calculation implementation for data queried from the
Germantown instance of JIRA.
"""

# Built-in modules
from collections import OrderedDict
from datetime import datetime
import heapq

# User-defined modules
from calculator import Calculator
from constants import *
from utilities import get_historical_dates, AverageAge

class JiraGTCalculator(Calculator):
  """
  Calculator used to calculate metrics for a given JIRA (Germantown) project.
  """
  
  def __init__(self, project, issuetype):
    """
    Initializes basic parameters regarding the metrics being calculated.
    
    @param project: Name of the project whose metrics are being calculated.
    @param issuetype: The issue type that the metrics should be associated with.
    """
    
    super(JiraGTCalculator, self).__init__(project, issuetype)
    
    # Sets the list of every possible status, in order
    self.status_list = [
       'New',                                                                       # 1 (1)
       'Approved', 'Open', 'Reopened', 'Deferred', 'In Progress',                   # 5 (6)
       'Change Ready', 'Change Required', 'In Review', 'Pending Approval',          # 4 (10)
       'In Dev', 'In Analysis', 'Dev Lead', 'Dev Lead Review',                      # 4 (14)
       'Build Ready', 'Test Ready', 'Passed to Test', 'Build Pending',              # 4 (18)
       'Test Build Pending', 'Deployed', 'Deployed to Test', 'Smoke Test',          # 4 (22)
       'In Test', 'Test Defects Pending', 'Test Lead Review',                       # 3 (25)
       'Passed to Prod', 'Prod Build Pending', 'Prod Defects Pending',              # 3 (28)
       'Resolved',                                                                  # 1 (29)
       'Closed'                                                                     # 1 (30)
    ]
    
    # Populates status group map
    self._set_status_group_map()
    
  def _set_status_group_map(self):
    """
    Sets the dictionary for the status map, mapping a status group to its given 
    statuses.
    """
    
    # Initializes the status list with New
    self.status_map['New'] = ['New']
    
    # Adds the Open status group for all associated issue types
    self.status_map['Open'] = []
    for index in range(1, 10):
      self.status_map['Open'].append(self.status_list[index])
        
    # Adds the In Dev status group for all associated issue types
    self.status_map['In Dev'] = []
    for index in range(10, 14):
      self.status_map['In Dev'].append(self.status_list[index])
        
    # Adds the Test status group for all associated issue types
    self.status_map['In Test'] = []
    for index in range(14, 25):
      self.status_map['In Test'].append(self.status_list[index])
        
    # Adds the Prod status group if the issue type is Defect
    self.status_map['In Prod'] = []
    for index in range(25, 28):
      self.status_map['In Prod'].append(self.status_list[index])
            
    # Adds the final two Resolved and Closed status groups
    for index in range(28, 30):
      self.status_map[self.status_list[index]] = [self.status_list[index]]
    
  def calculate(self, data, *args, **kwargs):
    """
    Calculates severity, status, and age data based on the given data passed.
    
    @param data: Data dictionary mapping issue key to its various associated 
    parameters.
    @return: A tuple with the following calculated data: (severity, status).
    """
    
    # Sets up priority queue, where data is prioritized by date
    queue = []
    
    # Sets up data dictionaries that will be used to contain calculated data
    severity_data = OrderedDict()
    status_data = OrderedDict()
    current_state = { }
    
    # List of fields used
    fields = [PROJECT, TRANS, STATUS, PRIORITY, COMPS, LINKS, PACK]
    
    # Populates priority queue with appropriate data
    for key, param_data in data.iteritems():
      # Grabs param_data fields
      priority = param_data[PRIORITY]
      created = param_data[CREATED]
      comps = param_data[COMPS]
      links = param_data[LINKS]
      pack = param_data[PACK]
      hist = param_data.get(HIST)
      proj = param_data.get(PROJECT, self.project)
      
      # Adds the first status (New) to the queue
      heapq.heappush(queue, (created, proj, key, 'New', priority, comps, links, pack))
      
      # Adds the historical statuses of the current JIRA item to the queue
      if (hist):
        for i, date in enumerate(hist[TRANS]):
          heapq.heappush(queue, (date, proj, key, hist[NEW][i], priority, comps, links, pack))
    
    # Iterates through dates to populate status and severity data dictionaries
    if (queue):
      earliest = queue[0][0]
      for date in get_historical_dates(earliest, self.extraction_day, False):
        # Pops items off queue until queue is empty or date limit is reached
        while(queue and queue[0][0].date() <= date):
          curr, proj, key, status, priority, comps, links, pack = heapq.heappop(queue)
          
          # Maps the key's current parameters, overwriting previous mapping
          current_state[key] = { }
          for field, value in zip(fields, [proj, curr, status, priority, comps, links, pack]):
            current_state[key][field] = value
          
        # Sets severity and status metric data at the given date
        severity_data[date] = self._get_severity_data(current_state)
        status_data[date] = self._get_status_data(current_state)
      
    # Gets age data separately from status and severity
    if ('age_data' not in kwargs or kwargs['age_data']):
      age_map = self._get_average_age_data(data)
    else: age_map = None
      
    return severity_data, status_data, age_map