"""
This is the module responsible for producing the presentation for ClearQuest.
"""

# Built-in modules
from collections import OrderedDict
from datetime import datetime

# User-defined modules
from calculator import DCRCalculator, RRCalculator, SCRCalculator
from constants import *
from directories import STATE_OF_QUALITY_DIR, CQ_DIR, GROUP_DIR, TMS_DIR, PROJECT_DIR
from powerpoint.data_source_ppt import DataSourcePPT
from utilities import create_dirpath

class ClearQuestPPT(DataSourcePPT):
  """
  This class produces the presentation for the Germantown instance Jira metrics.
  """
  
  def __init__(self, project_map):
    """
    Initializes the initial parameters for JIRA.
    
    @param project_map: Data dictionary mapping project groups to a list of tuples 
    containing associated project keys and names.
    """
    
    super(ClearQuestPPT, self).__init__(CQ_DIR)
    
    # Sets project group map
    self.project_map = project_map
    
    # Sets issue type list
    self.issue_types = OrderedDict([
      (DCR, [ENG_CHANGE, ENG_NOTICE]), (RR, [DEV, PROD]), (SCR, [DEFECT, ENHANCE])
    ])
    
    # Sets file paths for directories containing project data
    dir_link = [STATE_OF_QUALITY_DIR, CQ_DIR]
    self.project_path = create_dirpath(subdirs=dir_link + [PROJECT_DIR])
    self.project_group_path = create_dirpath(subdirs=dir_link + [GROUP_DIR])
    
  @staticmethod
  def _change_status_map(status_map):
    """
    Changes the structure of the given status map so that status groups are
    mapped to strings rather than lists of statuses.
    
    @param status_map: Data dictionary mapping status groups to lists of 
    associated statuses.
    @return: Data dictionary mapping status groups to strings of associated
    statuses.
    """
    
    new_map = OrderedDict()
    for status_group, status_list in status_map.iteritems():
      # Statuses are set differently based on status count
      if (len(status_list) == 1):
        new_map[status_group] = status_list[0]
      elif (len(status_list) == 2):
        new_map[status_group] = '%s & %s' % (status_list[0], status_list[1])
      else:
        new_map[status_group] = ', '.join(status_list[:-1])
        new_map[status_group] += ', & %s' % status_list[-1]
        
    return new_map
    
  def _add_things_to_know_slides(self):
    """
    Adds the two "Things to Know" slides to the presentation, one for the basic 
    information page, and the other one for the table of priorities.
    """
    
    # Sets bullet points
    bullets = [
      """This report is a snapshot of all active projects that TMS resources are working \
on, which use ClearQuest. The report is broken up and organized in alphabetical order, so you \
can use the CTRL-F feature to quickly locate what project(s) you are looking for.""",
      "The Average Aging tables refer to how many days on average do issues remain in a given status.",
      "The statuses used within the presentation are actually status groups, which may include numerous statuses.",
      """For any projects or project groups whose issues of a given issue type \
are all closed, all charts and tables related to that given issue type are excluded.""",
      "Only SCRs contain priorities."
    ]
    
    # Adds first Things to Know slide
    self.add_bullets_slide('Things to Know', bullets, log='Things to know slide added.')
    
    # Adds second Things to Know slide
    self.add_priority_table_slide('ClearQuest')
    
    # Sets up status group maps
    calc_tuples = [
      (DCRCalculator, DCR, ENG_CHANGE), (DCRCalculator, DCR, ENG_NOTICE), 
      (RRCalculator, RR, DEV), (RRCalculator, RR, PROD), (SCRCalculator, SCR, None)
    ]
    # Adds five status group slides
    for calc, metric_type, issue_type in calc_tuples:
      # Gets status map
      status_map = self._change_status_map(calc.get_status_group_map(issue_type))
      
      # Sets issue type text
      issue_text = metric_type
      if (issue_type): issue_text += ' - %s' % issue_type
      
      # Adds status table slide with given info
      self.add_status_table_slide(issue_text, status_map)
    
  def _add_regular_project_group_aging_slides(self, project_group):
    """
    Adds the aging slides for a regular project group.
    
    @param project_group: Name of project group associated with aging slides.
    """
    
    # Sets header map
    header_map = { }
    for metric_type, issue_types in self.issue_types.iteritems():
      for issue_type in issue_types:
        header = 'Overall "State of Quality": %s %ss' % (metric_type, issue_type)
        header_map[issue_type] = header
    
    # Adds aging slides for given project group
    self.add_project_group_aging_slides(project_group, header_map)
    
  def generate_presentation(self, date=datetime.today()):
    """
    Generates the full Powerpoint presentation for the Germantown instance of JIRA.
    
    @param date: Datetime object representing the date of the presentation, 
    defaulting to the current date.
    @return: File path of the generated Powerpoint presentation.
    """
    
    # Setups up the parameters for using the win32 functionalities
    self.setup()
    
    # Performs work for generating Powerpoint and cleans up afterward
    try:
      # Add title slide
      slide = self.add_title_slide(self.title, date.strftime('%B %d, %Y'))
      info = "Quality Assurance for Service Delivery\nGTS Technology Delivery Center"
      self.add_textframe(slide, info, 15, 0, False, Orientation=1, Left=20, 
                         Top=420, Width=500, Height=50)
      
      # Adds things to know slides
      self._add_things_to_know_slides()
      
      # Adds project group aging slides
      self._add_regular_project_group_aging_slides(TMS_DIR)
        
      # Adds aging, severity, and status slides for each project
      for key, name in self.project_map[TMS_DIR]:
        # Adds charts in the order of issue types
        for metric_type, issue_types in self.issue_types.iteritems():
          for issue_type in issue_types:
            issue_text = '%s %s' % (metric_type, issue_type)
            age_text = '%s Average Age Data' % issue_text
            self.add_project_aging_slides(key, name, issue_type, header=age_text, 
                                          keyword=metric_type)
            self.add_project_chart_slides(key, name, issue_text, SEV)
            self.add_project_chart_slides(key, name, issue_text, STATUS)
      
      # Adds concluding Xerox slide
      self.add_ending_slide()
      
      # Saves file
      save_path = '%s\\%s %s (ClearQuest).pptx' % (self.save_path, self.file_name, 
                                                date.strftime('%Y-%m-%d'))
      self.ppt_file.SaveAs(save_path)
    except Exception, e:
      print str(e)    # For troubleshooting
    #finally:
    self.cleanup()