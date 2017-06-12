"""
This module contains the calculation implementation for data queried from 
ClearQuest.
"""

# Built-in modules
from collections import OrderedDict
from datetime import datetime
import heapq

# User-defined modules
from calculator import Calculator
from constants import *
from utilities import get_historical_dates, AverageAge

class ClearQuestCalculator(Calculator):
  """
  This class encapsulates the code responsible for performing calculations
  on data from ClearQuest.
  """
  
  def __init__(self, project, issuetype):
    """
    Initializes basic parameters regarding the metrics being calculated.
    
    @param project: Name of the project whose metrics are being calculated.
    @param issuetype: The issue type that the metrics should be associated with.
    """
    
    super(ClearQuestCalculator, self).__init__(project, issuetype)
    
  @staticmethod
  def get_status_group_map(issuetype):
    """
    Gets the dictionary for the status map, mapping a status group to its given 
    statuses.
    
    @param issuetype: Issue type associated with the calculator.
    @return: Data dictionary for the status group mapping, based on the issue 
    type passed.
    """
    
    raise NotImplemented('Needs to be implemented by subclass.')
    
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
    fields = [PROJECT, TRANS, STATUS, PRIORITY]
    
    # Populates priority queue with appropriate data
    for key, param_data in data.iteritems():
      # Grabs param_data fields
      priority = param_data.get(PRIORITY, None)
      hist = param_data.get(HIST, None)
      proj = param_data.get(PROJECT, self.project)
      
      # Adds the historical statuses of the current JIRA item to the queue
      if (hist):
        for i, date in enumerate(hist[TRANS]):
          heapq.heappush(queue, (date, proj, key, hist[NEW][i], priority))
    
    # Iterates through dates to populate status and severity data dictionaries
    if (queue):
      earliest = queue[0][0]
      for date in get_historical_dates(earliest, self.extraction_day, False):
        # Pops items off queue until queue is empty or date limit is reached
        while(queue and queue[0][0].date() <= date):
          curr, proj, key, status, priority = heapq.heappop(queue)
          
          # Maps the key's current parameters, overwriting previous mapping
          current_state[key] = { }
          for field, value in zip(fields, [proj, curr, status, priority]):
            current_state[key][field] = value
          
        # Sets severity and status metric data at the given date
        severity_data[date] = self._get_severity_data(current_state)
        status_data[date] = self._get_status_data(current_state)
      
    # Gets age data separately from status and severity
    age_map = self._get_average_age_data(data)
      
    return severity_data, status_data, age_map
  
# Imports ClearQuest metric types to top level
from dcr import DCRCalculator
from rr import RRCalculator
from scr import SCRCalculator