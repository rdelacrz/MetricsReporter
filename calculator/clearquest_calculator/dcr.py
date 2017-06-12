"""
This module contains the calculation implementation for DCR data queried from 
ClearQuest.
"""

# Built-in modules
from collections import OrderedDict

# User-defined modules
from constants import ENG_CHANGE, ENG_NOTICE
from calculator.clearquest_calculator import ClearQuestCalculator

class DCRCalculator(ClearQuestCalculator):
  """
  This class encapsulates the code responsible for DCR calculations.
  """
  
  def __init__(self, project, issuetype):
    """
    Initializes basic parameters regarding the DCR data being calculated.
    
    @param project: Name of the project whose metrics are being calculated.
    @param issuetype: The issue type that the metrics should be associated with.
    """
    
    super(DCRCalculator, self).__init__(project, issuetype)
    
    # DCRs have no priority data
    self.priority_list = []
    
    # Sets up DCR status buckets
    self.status_map = self.get_status_group_map(issuetype)
      
  @staticmethod
  def get_status_group_map(issuetype):
    """
    Gets the dictionary for the status map, mapping a status group to its given 
    statuses.
    
    @param issuetype: Issue type associated with the calculator.
    @return: Data dictionary for the status group mapping, based on the issue 
    type passed.
    """
    
    if (issuetype == ENG_CHANGE):
      return OrderedDict([
        ('Submitted', ['Submitted']),
        ('In Review', ['In Review', 'Review Complete']),
        ('Ready For Release', ['Ready For Release']),
        ('Void', ['Void']),
        ('Closed', ['Closed'])
      ])
    elif (issuetype == ENG_NOTICE):
      return OrderedDict([
        ('Submitted', ['Submitted']),
        ('ID Generated', ['ID Generated']),
        ('In Review', ['Ready For Review', 'In Review', 'Review Complete']),
        ('Ready For Release', ['Ready For Release']),
        ('Void', ['Void']),
        ('Closed', ['Closed'])
      ])
    else:
      return OrderedDict()