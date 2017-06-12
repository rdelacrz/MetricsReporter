"""
This is the module responsible for producing the presentation for the 
Germantown instance of JIRA.
"""

# Built-in modules
from datetime import datetime

# User-defined modules
from constants import CR, DEFECT, TASK, BLOCKER, CRITICAL, MINOR, MAJOR, \
                      TRIVIAL, SEV, STATUS
from directories import GT_JIRA_DIR, STATE_OF_QUALITY_DIR, GROUP_DIR, \
                        PROJECT_DIR, TDC_DIR
from powerpoint.data_source_ppt import DataSourcePPT
from utilities import create_dirpath

# Sets aging slide note for TDC-wide slides (added below the aging table)
TDC_AGING_NOTES = """* The Overall column for statuses corresponds to the average \
overall age of all the issues. The age of each individual issue is equal to the \
amount of time that has elapsed between the issue's creation date and closure date. \
If an issue is not closed, the current date is used instead of closure date.
* The Overall row for priorities corresponds to all issues within a given status \
group, regardless of what priority it has."""

class GTJiraPPT(DataSourcePPT):
  """
  This class produces the presentation for the Germantown instance Jira metrics.
  """
  
  def __init__(self, project_map):
    """
    Initializes the initial parameters for JIRA.
    
    @param project_map: Data dictionary mapping project groups to a list of tuples 
    containing associated project keys and names.
    """
    
    super(GTJiraPPT, self).__init__(GT_JIRA_DIR)
    
    # Sets project group map
    self.project_map = project_map
    
    # Sets issue type list
    self.issue_types = [CR, DEFECT, TASK]
    
    # Sets issue type definitions
    self.issue_def = {
      CR : 'A CR is a request for changes to the project solution that are' + \
          ' submitted for analysis, evaluation, assignment, and resolution.',
      DEFECT : 'A Defect is the result of a nonconformity discovered during' + \
              ' testing. It follows an identical workflow to CRs and are' + \
              ' linked to CRs where applicable.',
      TASK : 'A Task is an issue in JIRA, which represents a question, ' + \
            'problem, or condition that requires a decision and resolution.'
    }
    
    # Sets file paths for directories containing project data
    dir_link = [STATE_OF_QUALITY_DIR, GT_JIRA_DIR]
    self.project_path = create_dirpath(subdirs=dir_link + [PROJECT_DIR])
    self.project_group_path = create_dirpath(subdirs=dir_link + [GROUP_DIR])
    
  def _add_things_to_know_slides(self):
    """
    Adds the two "Things to Know" slides to the presentation, one for the basic 
    information page, and the other one for the table of priorities.
    """
    
    # Sets bullet points
    bullets = [
      """This report is a snapshot of all active projects that TDC resources are working \
on, which use JIRA.  The report is broken up and organized in alphabetical order, so you \
can use the CTRL-F feature to quickly locate what project(s) you are looking for.""",
      "The Average Aging tables refer to how many days on average do issues remain in a given status.",
      "The statuses used within the presentation are actually status groups, which may include numerous statuses.",
      """For any projects or project groups whose issues of a given issue type \
are all closed, all charts and tables related to that given issue type are excluded.""",
      [
        """"Open" includes Approved, Open, Reopened, Deferred, In Progress, Change Ready, \
Change Required, In Review, & Pending Approval.""",
        "\"In Dev\" includes In Dev, In Analysis, Dev Lead, & Dev Lead Review.",
        """"In Test" includes Build Ready, Test Ready, Passed to Test, Build Pending, Test \
Build Pending, Deployed, Deployed to Test, Smoke Test, In Test, Test Defects Pending, & Test Lead Review.""",
        "\"In Prod\" includes Passed to Prod, Prod Build Pending, & Prod Defects Pending."
      ]
    ]
    
    # Adds first Things to Know slide
    self.add_bullets_slide('Things to Know', bullets, log='Things to know slide added.')
    
    # Adds second Things to Know slide
    self.add_priority_table_slide('JIRA')
    
  def _add_tdc_wide_aging_slides(self):
    """
    Adds the State of Quality aging slides for all of the TDC.
    """
    
    # Sets header (Change Request shortened to CR)
    header_map = { }
    for issue_type in self.issue_types:
      if (issue_type == CR): header = '"State of Quality" of TDC: CRs'
      else: header = '"State of Quality" of TDC: %ss' % issue_type
      header_map[issue_type] = header
    
    # Adds TDC wide aging slides
    self.add_project_group_aging_slides(TDC_DIR, header_map, note=TDC_AGING_NOTES, 
                                        write_group=False)
    
  def _add_regular_project_group_aging_slides(self, project_group):
    """
    Adds the aging slides for a regular project group (rather than TDC wide).
    
    @param project_group: Name of project group associated with aging slides.
    """
    
    # Sets header (Change Request shortened to CR)
    header_map = { }
    for issue_type in self.issue_types:
      if (issue_type == CR): header = 'Overall "State of Quality": CRs'
      else: header = 'Overall "State of Quality": %ss' % issue_type
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
      
      # Adds State of Quality overall aging slides
      self._add_tdc_wide_aging_slides()
      
      # Prepares header map for Status and Severity charts (only puts CR in it)
      header_map = { CR : { SEV: 'CR Priority Data (Cumulative View)', 
                           STATUS : 'CR Workflow Activity (Cumulative View)'}}
      
      # Adds slides for each project group
      for project_group, project_list in self.project_map.iteritems():
        # Adds project group intro slide
        self.add_project_group_slide(project_group)
        
        # Adds project group aging slides
        self._add_regular_project_group_aging_slides(project_group)
        
        # Adds aging, severity, and status slides for each project
        for key, name in project_list:
          # Adds charts in the order of issue types
          for issue_type in self.issue_types:
            self.add_project_aging_slides(key, name, issue_type)
            self.add_project_chart_slides(key, name, issue_type, SEV, header_map)
            self.add_project_chart_slides(key, name, issue_type, STATUS, header_map)
      
      # Adds individual priority slides
      for priority in [BLOCKER, CRITICAL, MAJOR, MINOR, TRIVIAL]:
        self.add_individual_priority_slide(priority)
      
      # Adds concluding Xerox slide
      self.add_ending_slide()
      
      # Saves file
      save_path = '%s\\%s %s (GT JIRA).pptx' % (self.save_path, self.file_name, 
                                                date.strftime('%Y-%m-%d'))
      self.ppt_file.SaveAs(save_path)
    finally:
      self.cleanup()