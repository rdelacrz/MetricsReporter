"""
This module is responsible for calculating metrics based on queried data that 
is passed.
"""

# Built-in modules
from collections import OrderedDict
from datetime import datetime

# User-defined modules
from constants import *
from utilities import AverageAge

class Calculator(object):
  """
  Default calculator used to calculate metrics for a given project.
  """
  
  def __init__(self, project, issuetype):
    """
    Initializes basic parameters regarding the metrics being calculated.
    
    @param project: Name of the project whose metrics are being calculated.
    @param issuetype: The issue type that the metrics should be associated with.
    """
    
    # Sets initial parameters
    self.project = project
    self.issuetype = issuetype
    
    # Week day for extraction (0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun)
    self.extraction_day = 4
    
    # Priority list
    self.priority_list = ['Blocker', 'Critical', 'Major' ,'Minor', 'Trivial']
    
    # Sets status group map
    self.status_map = OrderedDict()
    
  def get_status_desc(self):
    """
    Gets a list of tuples, matching each status group to its corresponding 
    description.
    
    @return: List of tuples with following format: (status group, description).
    """
    
    desc_list = []
    
    for status_group, status_list in self.status_map.iteritems():
      # Writes desc differently based on status list length
      if (len(status_list) == 1):
        desc = status_list[0]
      elif (len(status_list) == 2):
        desc = '%s includes %s & %s' % (status_group, status_list[0], status_list[1])
      else:
        desc = '%s includes ' % status_group
        desc += ', '.join(status_list[:-1])
        desc += ' & %s' % status_list[-1]
      desc_list.append((status_group, desc))
      
    return desc_list
      
  def _get_status_group(self, status):
    """
    Returns the status group of the given status.
    
    @param status: Status whose status group is being searched for.
    @return: The name of the status group that the status is part of.
    """
    
    # Iterates through each of the available status groups
    status = ' '.join(status.split()).strip()
    for status_group, status_list in self.status_map.iteritems():
      if (status in status_list): return status_group
    return None
    
  def _get_severity_data(self, data):
    """
    Parses the given data and return a data dictionary for the priority counts
    within the data the following mapping:
    
    [TOTAL, OPEN, CLOSED] -> <int counts>
    
    @param data: A data dictionary containing the current states of the issues.
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
            if (status in self.status_map[CLOSED]): severity_data[CLOSED][p] += 1
            else:                  severity_data[OPEN][p] += 1
            severity_data[TOTAL][p] += 1
    
    return severity_data
  
  def _get_status_data(self, data):
    """
    Parses the given data and return a data dictionary for the status counts 
    within the data the following mapping:
    
    [TOTAL] -> <int count>
    
    @param data: A data dictionary containing the current states of the issues.
    """
      
    # Initializes data dictionary named status_data
    data_types = [TOTAL]
    status_data = OrderedDict([(x, OrderedDict([
       (group, 0) for group in self.status_map.keys() + [TOTAL]                                      
    ])) for x in data_types])
    
    # Iterates through each issue
    for param in data.values():
      status = param[STATUS]
      
      # Formats status
      status = self._get_status_group(status)
              
      # Increments status counts on both the status and Total level
      if (status):
        for s in [status, TOTAL]:
          status_data[TOTAL][s] += 1
    
    return status_data
  
  def _set_age_map_value(self, age_map, priority, status_group, upper_date, lower_date):
    """
    Sets the age map values based on the given parameters of the current issue.
    
    @param age_map: Data dictionary containing averages, segmented by priority and 
    status group.
    @param priority: Priority of the current issue.
    @param status_group: Status group under which the current issue's status is part of.
    @param upper_date: Marks the end of a status group.
    @param lower_date: Marks the beginning of a status group.
    """
    
    # Gets time difference between the current status and previous status group
    diff = upper_date - lower_date
          
    # Adds age information for Overall as well
    p_list = [priority, OVERALL] if (priority) else [OVERALL]
    for p in p_list:
      if (status_group in age_map[p]):
        age_map[p][status_group].update(diff, update_avg=False)
  
  def _get_average_age_data(self, data):
    """
    Calculates average age (in number of days) of all the issues based on the 
    data passed, broken down by priority and status.
    
    @param data: A data dictionary containing the parameters and history of
    each parameter.
    @return: A data dictionary containing AverageAge objects containing the 
    average age of the issues in a given priority/status group. It has the
    following structure:
      <priority> -> <status group> -> AverageAge object
    """
    
    # Sets the status and priority lists (appended with Overall)
    status_list = self.status_map.keys() + [OVERALL]
    priority_list = list(self.priority_list) + [OVERALL]
    
    # Initializes the age map
    age_map = OrderedDict([
      (priority, OrderedDict([
        (status, AverageAge()) for status in status_list
      ])) for priority in priority_list
    ])
    
    # Iterates through each issue
    for param_data in data.values():
      # Grabs param_data fields relevant to age average calculations
      priority = param_data.get(PRIORITY)
      status = param_data[STATUS]
      created = param_data.get(CREATED, param_data.get(SUBMIT_DATE))
      hist = param_data.get(HIST)
      
      # Sets status history tuple list
      if (hist):
        # Accounts for change histories that include the very first issues
        if (hist[OLD][0]):
          first_status = hist[OLD][0]
          history_list = [(first_status, created)] + zip(hist[NEW], hist[TRANS])
        else:
          first_status = hist[NEW][0]
          history_list = zip(hist[NEW], hist[TRANS])
        curr_group = self._get_status_group(first_status)
      else:
        history_list = [(status, created)]
        curr_group = self._get_status_group(status)
        
      # Sets current date of status group
      curr_date = created
        
      # Iterates through history list
      for status, status_date in history_list:
        # Gets the status group of the given status
        group = self._get_status_group(status)
      
        # Marks status group change
        if (group != curr_group):
          self._set_age_map_value(age_map, priority, curr_group, status_date, curr_date)
          
          # Sets current status parameters
          curr_group = group
          curr_date = status_date
          
      # Determines age of current status group
      curr_date = datetime.today()
      self._set_age_map_value(age_map, priority, group, curr_date, status_date)
      
      # Determines Overall age of an issue (either to Closed or to current date)
      if (curr_group == CLOSED):
        self._set_age_map_value(age_map, priority, OVERALL, status_date, created)
      else:
        self._set_age_map_value(age_map, priority, OVERALL, curr_date, created)
        
    # Performs calculations
    for status_map in age_map.values(): 
      for average in status_map.values(): average.calculate_average()
        
    return age_map
  
  @staticmethod
  def combine_age_data(data, issue_type):
    """
    Takes all the age data dictionaries from the values of the top-level data
    dictionary (which are being mapped to by projects or project groups), and 
    combines them into one data dictionary that determines the average ages of 
    every associated project or group (of a given issue type).
    
    @param data: Multi-level data dictionary that maps projects or project
    project groups to data dictionaries containing age data. It has the
    following structure:
      <project/group name> -> <issue_type> -> <priority> -> <status group> 
        -> AverageAge object
    @param issue_type: The issue type intended for age calculation.
    @return: Multi-level data dictionary that combines all the lower-level data 
    dictionaries of the one passed as a parameter into one consolidated one,
    containing the average age data for all of the associated projects/project 
    groups as a whole.
    """
    
    combined_data = OrderedDict()
    
    # Iterates through every project's or group's age data
    for age_data in data.values():
      for priority, status_data in age_data[issue_type].iteritems():
        # Maps priority if it currently doesn't exist
        if (priority not in combined_data):
          combined_data[priority] = OrderedDict()
        for status_group, age_obj in status_data.iteritems():
          # Maps status group if it currently doesn't exist
          if (status_group not in combined_data[priority]):
            combined_data[priority][status_group] = AverageAge()
          combined_data[priority][status_group].combine(age_obj)
    
    return combined_data
    
  def calculate(self, data, *args, **kwargs):
    """
    Calculates metrics data based on the given data passed.
    """
    
    raise NotImplemented("calculate() function needs to be implemented by a sub-class")
  
  def get_metrics(self, data):
    """
    Performs calculation on given data and returns the resulting calculation.
    """
    
    print "Performing calculations for %s data..." % self.project
    return self.calculate(data)
  
# Brings children classes to top level
from clearquest_calculator import DCRCalculator, RRCalculator, SCRCalculator
from jira_gt_calculator import JiraGTCalculator
from jira_gt_calculator.septa import SEPTACalculator
from jira_gt_calculator.compliance import ComplianceCalculator
      