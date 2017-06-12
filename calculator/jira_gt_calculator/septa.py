"""
A specialized module specifically for the SEPTA project.
"""

# Built-in modules
from collections import OrderedDict

# User-defined modules
from constants import *
from calculator.jira_gt_calculator import JiraGTCalculator

class SEPTACalculator(JiraGTCalculator):
  """
  Contains the code for specifically calculating SEPTA's metrics.
  """
  
  def __init__(self):
    """
    Initializes the parameters specific to SEPTA.
    """
    
    super(SEPTACalculator, self).__init__('SEPTA', DEFECT)
    
    # Sets extraction day to Saturday
    self.extraction_day = 5
    
  def _set_status_group_map(self):
    """
    Sets the dictionary for the status map, mapping a status group to its given 
    statuses.
    """
    
    # Initializes the status map for various parameters
    self.status_map = OrderedDict()
    self.status_map['In Dev'] = ['New', 'In Dev', 'In Analysis', 'Dev Lead Review', 
                                 'Test Defects Pending', 'Prod Defects Pending']
    self.status_map['Passed to Test'] = ['Passed to Test']
    self.status_map['Test Build Pending'] = ['Test Build Pending']
        
    # Adds the Test status group for certain statuses
    self.status_map['In Test'] = ['Deployed to Test', 'In Test','Test Lead Review']
            
    # Adds the Closed status group
    self.status_map['Closed'] = ['Passed to Prod', 'Prod Build Pending', 'Resolved', 'Closed']
    
  def _get_severity_data(self, data):
    """
    Parses the given data and return a data dictionary for the priority 
    counts within the data the following mapping:
    
    [TOTAL, OPEN, CLOSED] -> <status> -> int count
    
    @param data: A data dictionary containing the current states of the
    issues.
    """
    
    # Initializes list of data types
    data_types = [TOTAL, CLOSED, OPEN]
    for original_type in [TOTAL, CLOSED, OPEN]:
      for sub_type in [FAT_A, FAT_B, PILOT, PILOT_HI]:
        data_types.append('%s (%s)' % (original_type, sub_type))
        
    # Initializes data dictionary named severity_data
    severity_data = OrderedDict([(x, OrderedDict([
       (priority, 0) for priority in self.priority_list + [TOTAL]                                      
    ])) for x in data_types])
    
    # Iterates through each issue
    for param in data.values():
      status = param[STATUS]
      priority = param[PRIORITY]
      comps = param[COMPS]
      linked = param[LINKS]
      pack = param[PACK]
          
      # Increments priority counts depending on closure
      if (priority):
        # Skips hardware
        if (comps is None or ('hardware' not in comps.lower() and 'hw' not in comps.lower()
                              and 'security' != comps.lower())):
          
          # Increments priority counts on both the priority and Total level
          for p in [priority, TOTAL]:
            # Regular severity counts
            if (status in self.status_map[CLOSED]): severity_data[CLOSED][p] += 1
            else:                                   severity_data[OPEN][p] += 1
            severity_data[TOTAL][p] += 1
            
            # Fat-A, Fat-B, Pilot
            for cond, data_type in [('PACK-151' in linked, FAT_A), 
                (FAT_B == pack, FAT_B), ('PILOT' == pack, PILOT),
                ('PILOT' == pack and priority not in ['Minor', 'Trivial'], PILOT_HI)]:
              if (cond):
                if (status in self.status_map[CLOSED]): 
                  severity_data['%s (%s)' % (CLOSED, data_type)][p] += 1
                else:                  
                  severity_data['%s (%s)' % (OPEN, data_type)][p] += 1
                severity_data['%s (%s)' % (TOTAL, data_type)][p] += 1
    
    # Removes Minor and Trivial categories from Hi Priority
    for data_type in [TOTAL, OPEN, CLOSED]:
      del severity_data['%s (%s)' % (data_type, PILOT_HI)]['Minor']
      del severity_data['%s (%s)' % (data_type, PILOT_HI)]['Trivial']
    
    return severity_data
  
  def _get_status_data(self, data):
    """
    Parses the given data and return a data dictionary for the status counts 
    within the data the following mapping:
    
    [FAT_1A, FAT_1B, FAT_1B_HI, PILOT, PILOT_HI, TOTAL] -> <status> -> int count
    
    @param data: A data dictionary containing the current states of the
    issues.
    """
      
    # Initializes data dictionary named status_data
    data_types = [FAT_1A, FAT_1B, FAT_1B_HI, PILOT, PILOT_HI, TOTAL]
    status_data = OrderedDict([(x, OrderedDict([
       (group, 0) for group in self.status_map.keys() + [TOTAL]                                      
    ])) for x in data_types])
    
    # Iterates through each issue
    for param in data.values():
      status = param[STATUS]
      priority = param[PRIORITY]
      comps = param[COMPS]
      linked = param[LINKS]
      pack = param[PACK]
      
      # Skips hardware
      if (comps is None or ('hardware' not in comps.lower() and 'hw' not in comps.lower()
                            and 'security' != comps.lower())):
        # Formats status
        status = self._get_status_group(status)
              
        # Increments status counts on both the status and Total level
        if (status):
          for s in [status, TOTAL]:
            # FAT-1A, FAT-1B, Hi Priority FAT-1B, Pilot, Hi Priority Pilot
            for cond, data_type in [('PACK-151' in linked, FAT_1A), 
                  ('FAT-B' == pack, FAT_1B), 
                  ('FAT-B' == pack and priority not in ['Minor', 'Trivial'], FAT_1B_HI),
                  ('PILOT' == pack, PILOT), 
                  ('PILOT' == pack and priority not in ['Minor', 'Trivial'], PILOT_HI)   
                ]:
              if (cond):
                status_data[data_type][s] += 1
              
            # Sets total count
            status_data[TOTAL][s] += 1
    
    return status_data