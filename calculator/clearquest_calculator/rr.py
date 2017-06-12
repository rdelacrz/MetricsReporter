"""
This module contains the calculation implementation for RR data queried from 
ClearQuest.
"""

# Built-in modules
from collections import OrderedDict

# User-defined modules
from constants import DEV, PROD
from calculator.clearquest_calculator import ClearQuestCalculator

class RRCalculator(ClearQuestCalculator):
  """
  This class encapsulates the code responsible for RR calculations.
  """
  
  def __init__(self, project, issuetype):
    """
    Initializes basic parameters regarding the RR data being calculated.
    
    @param project: Name of the project whose metrics are being calculated.
    @param issuetype: The issue type that the metrics should be associated with.
    """
    
    super(RRCalculator, self).__init__(project, issuetype)
    
    # RRs have no priority data
    self.priority_list = []
    
    # Sets up RR status buckets based on issue type given
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
    
    if (issuetype == DEV):
      return OrderedDict([
        ('Submitted', ['Submitted']),
        ('Dev Release Approved', ['Dev Release Approved']),
        ('In Build', ['Waiting To Build', 'Build Failed', 'Build Approved']),
        ('In Engineering', ['Engineering Test', 'Engineering Failed']),
        ('In Release', ['Ready For Release', 'Dev Release Failed', 'Dev Release Passed']),
        ('Closed', ['Closed'])
      ])
    elif (issuetype == PROD):
      return OrderedDict([
        ('Submitted', ['Submitted']),
        ('Program Approved', ['Program Approved']),
        ('In Review', ['Engg Reviewed', 'CCB Approved', 'CCB Rejected', 'Patch CM', 'Patch Reviewed']),
        ('In Build', ['Waiting To Build', 'Build Failed', 'Build Approved']),
        ('In Engineering', ['Engineering Test', 'Engineering Failed', 'Ready For Release']),
        ('In System Testing', ['System Testing', 'System Test Failed']),
        ('In Fielding', ['Ready To Field', 'Field Test In Progress', 'Field Test Failed', 'Fielded']),
        ('SE Approved', ['SE Reviewed', 'SE Approved']),
        ('Cancelled', ['Release Withdrawn', 'Cancelled']),
        ('Closed', ['Closed'])
      ])
    else:
      return OrderedDict()   # This shouldn't be reached