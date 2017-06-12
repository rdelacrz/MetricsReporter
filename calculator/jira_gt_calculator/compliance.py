"""
A specialized module specifically for compliance issues.
"""

# Built-in modules
from collections import OrderedDict

# User-defined modules
from constants import *
from calculator.jira_gt_calculator import JiraGTCalculator
from db_accessor.jira_gt import JiraGT

class ComplianceCalculator(JiraGTCalculator):
  """
  Calculator used to calculate metrics for a given JIRA (Germantown) project.
  """
  
  def __init__(self):
    """
    Initializes basic parameters regarding the metrics being calculated.
    
    @param project: Name of the project whose metrics are being calculated.
    """
    
    super(ComplianceCalculator, self).__init__('Compliance', 'Compliance')
    
    # Determines list of projects with Compliance issues
    self.project_list = self._get_compliance_projects()
    
  def _get_compliance_projects(self):
    """
    Performs Oracle query to get a list of current projects with the Compliance
    issue type.
    """
    
    # Sets query string
    query = """
      SELECT DISTINCT project.pkey FROM jiraissue
        JOIN project   ON project.id=jiraissue.project
        JOIN issuetype ON issuetype.id=jiraissue.issuetype
      WHERE issuetype.pname='Compliance' ORDER BY project.pkey
    """
    
    # Connects to JIRA backend and gets Compliance projects
    jira = JiraGT()
    jira.connect()
    data = jira.query(query)
    jira.disconnect()
    
    return [project[0] for project in data]
    
  def _get_severity_data(self, data):
    """
    Parses the given data and return a data dictionary for the priority 
    counts within the data the following mapping:
    
    [TOTAL, OPEN, CLOSED] -> [int counts]
    
    @param data: A data dictionary containing the current states of the
    issues.
    """
        
    # Initializes data dictionary named severity_data
    data_types = [TOTAL, CLOSED, OPEN]
    severity_data = OrderedDict([(x, OrderedDict([
       (priority, 0) for priority in self.priority_list + [TOTAL]                                      
    ])) for x in data_types])
    
    # Iterates through each issue
    for param in data.values():
      status = param[STATUS]
      priority = param[PRIORITY]
          
      # Increments priority counts depending on closure
      if (priority):
          # Increments priority counts on both the priority and Total level
          for p in [priority, TOTAL]:
            # Regular severity counts
            if (status == CLOSED): severity_data[CLOSED][p] += 1
            else:                  severity_data[OPEN][p] += 1
            severity_data[TOTAL][p] += 1
    
    return severity_data
  
  def _get_status_data(self, data):
    """
    Parses the given data and return a data dictionary for the status counts 
    within the data the following mapping:
    
    [TOTAL, projects] -> [int count]
    
    @param data: A data dictionary containing the current states of the
    issues.
    """
      
    # Initializes data dictionary named status_data
    data_types = [TOTAL] + self.project_list
    status_data = OrderedDict([(x, OrderedDict([
       (group, 0) for group in self.status_map.keys() + [TOTAL]                                      
    ])) for x in data_types])
    
    # Iterates through each issue
    for param in data.values():
      status = param[STATUS]
      project = param[PROJECT]
      
      # Formats status
      status = self._get_status_group(status)
              
      # Increments status counts on both the status and Total level
      if (status):
        for s in [status, TOTAL]:
          status_data[project][s] += 1
          status_data[TOTAL][s] += 1
    
    return status_data