"""
This module contains the calculation implementation for SCR data queried from 
ClearQuest.
"""

# Built-in modules
from collections import OrderedDict

# User-defined modules
from calculator.clearquest_calculator import ClearQuestCalculator

class SCRCalculator(ClearQuestCalculator):
  """
  This class encapsulates the code responsible for SCR calculations.
  """
  
  def __init__(self, project, issuetype):
    """
    Initializes basic parameters regarding the SCR data being calculated.
    
    @param project: Name of the project whose metrics are being calculated.
    @param issuetype: The issue type that the metrics should be associated with.
    """
    
    super(SCRCalculator, self).__init__(project, issuetype)
    
    # Sets up SCR status buckets
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
    
    return OrderedDict([
      ('New', ['New']),
      ('Open', ['Assigned', 'Open', 'Wait', 'Approved', 'Resolved', 'Verified', 'Merged', 'Postponed', 'Duplicate']),
      ('Document Impacted', ['Document Impacted']),
      ('In Test', ['Baselined', 'System Tested', 'Ready To Field', 'Fielded']),
      ('SE Reviewed', ['SE Reviewed', 'SE Approved']),
      ('Closed', ['Closed'])
    ])