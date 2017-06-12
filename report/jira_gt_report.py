"""
Produces report specifically for the Germantown instance of JIRA.
"""

# Built-in modules
from collections import OrderedDict

# User-defined modules
from calculator import JiraGTCalculator
from constants import *
from directories import PROJECT_DIR, MATRICES_DIR, GT_JIRA_DIR, GROUP_DIR, TDC_DIR
from db_accessor import JiraGT
from powerpoint import GTJiraPPT
from report import Report
from utilities import create_dirpath
import xlrd

# Set to determine whether project groupings are pulled from Excel or not
USE_EXCEL = True

class JiraGTReport(Report):
  """
  Encapsulates the reporting code specifically for the Germantown instance of 
  JIRA.
  """
  
  def __init__(self, project_map=None):
    """
    Initializes JIRA-specific parameters.
    
    @param project_map: A customizable project map mapping project group to
    lists of associated projects. If this parameter is left blank, the default
    project mapping obtained by self._set_project_map() will be used instead.
    """
    
    # Initializes initial parameters
    super(JiraGTReport, self).__init__()
    
    # Sets database to Jira instance
    self.db = JiraGT()
    
    # Sets calculator to Jira calculator
    self.calc = JiraGTCalculator
    
    # Sets project map
    self.project_map = project_map if (project_map) else self._get_project_map()
    
    # Sets Powerpoint generator
    self.ppt = GTJiraPPT(self.project_map)
    
    # Data source name associated with the report
    self.data_source = 'Jira (Germantown)'
    
    # Appends to base trail
    self.base_dir_trail += [GT_JIRA_DIR]
    
    # Issue types associated with the report
    self.issue_types = [CR, DEFECT, TASK]
    
    # Raw data headers (a list of list of tuples with header name & cell width)
    self.raw_data_headers = [
      [(KEY, 12), (ISSUETYPE, 16), (PRIORITY, 10), ('Current Status', 16), 
              (CREATED, 17), (RESOLVED, 17), (COMPS, 17), (LINKS, 25), (PACK, 12), 
              (DEV_EST, 14), (DEV_ACT, 14)],
      [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 17)]
    ]
    
  def _read_excel_project_map(self):
    """
    Reads the group to project mappings from the JIRA Projects Excel file.
    
    @return: Data dictionary mapping project group names to a list of tuples
    with project keys and associated full project names. It looks like:
      <project group> -> [(project key, full project name),]
    """
    
    # Initializes group map
    group_map = OrderedDict()
    
    # File path of JIRA projects file
    filepath = '%s\\JIRA Projects.xlsx' % create_dirpath(subdirs=[MATRICES_DIR])
    
    # Workbook for file
    wb = xlrd.open_workbook(filepath)
    
    # Sheet with project information
    sheet = wb.sheet_by_index(0)
    
    # Traverses through each row for projects.
    for row in range(sheet.nrows)[1:]:
      name = str(sheet.cell(row, 0).value).strip()
      pkey = str(sheet.cell(row, 1).value).strip()
      group = str(sheet.cell(row, 5).value).strip()
      active = str(sheet.cell(row, 6).value).strip()
      if (group not in ['', 'n/a'] and active not in ['no', 'n/a']):
        if (group not in group_map):
          group_map[group] = []
        group_map[group].append((pkey, name))
    
    # Releases resources using the file to restore used memory
    wb.release_resources()
    
    return group_map
    
  def _get_project_map(self):
    """
    Gets the project group to project list mapping for the current report. 
    
    @return: Data dictionary mapping project group names to a list of tuples
    with project keys and associated full project names. It looks like:
      <project group> -> [(project key, full project name),]
    """
    
    if (USE_EXCEL): return self._read_excel_project_map()
    else:           return self.db.get_active_projects()
    
  def _query_data(self, *args, **kwargs):
    """
    Iteratively queries data from the data source.
    
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    @yield: A tuple of three items, with the first two items being group name
    and project name, and the third item being a data dictionary containing the
    associated queried data for raw data files and metric Excel reports.
    """
    
    # Iterates through the projects within each project group
    for group, projects in self.project_map.iteritems():
      for proj, _ in projects:
        data = self.db.get_issue_data(proj, issuetype=self.issue_types)
        yield group, proj, data
    
  def _produce_metric_files(self, project, data, *args, **kwargs):
    """
    Produces all the Excel metric files to be used in the final Powerpoint 
    report.
    
    @param project: Name of project for which metric files are being produced.
    @param data: Data from which calculated metrics are derived.
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    @return: Tuple of three data dictionaries for severity data, status data,
    and age data.
    """
    
    # Initializes segmented data dictionary
    segmented_data = OrderedDict([(x, OrderedDict()) for x in self.issue_types])
    
    # Segments data by issue type
    for key, params in data.iteritems():
      segmented_data[params[ISSUETYPE]][key] = params
      
    # Performs calculations for each issue type
    sev_data = OrderedDict()
    status_data = OrderedDict()
    age_data = OrderedDict()
    for issue_type in self.issue_types:
      calc = self.calc(project, issue_type)
      sev, status, age = calc.get_metrics(segmented_data[issue_type])
      
      # Inserts data into data dictionaries
      sev_data[issue_type] = sev
      status_data[issue_type] = status
      age_data[issue_type] = age
      
      # Produces severity file
      series_names = [(p, p) for p in calc.priority_list]
      self._produce_chart_file(project, sev, issue_type, SEV, series_names)
      
      # Produces status file
      series_names = calc.get_status_desc()
      self._produce_chart_file(project, status, issue_type, STATUS, series_names)
      
    # Produces age file
    self._produce_age_file(age_data, self.base_dir_trail + [PROJECT_DIR, project])
    
    return sev_data, status_data, age_data
  
  def _produce_group_age_tables(self, data, *args, **kwargs):
    """
    Combines data dictionary age data of every group contained inside, and 
    outputs Average Age data file for all of them. Afterwards, produces an
    overall Average Age data file for all contained projects.
    
    @param data: Multi-level data dictionary mapping groups to projects to age
    data dictionaries. It has the following structure:
      <project group> -> <project> -> <issue type> -> <priority> 
        -> <status group> -> <AverageAge object>
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    """
    
    # Initializes project group level aging data
    group_age = OrderedDict([(group, OrderedDict()) for group in data])
    
    # Combines data among each project group and produces age data files
    for group, age_data in data.iteritems():
      for issue in self.issue_types:
        group_age[group][issue] = self.calc.combine_age_data(age_data, issue)
        
      # Creates age data file for project group
      save_path_trail = self.base_dir_trail + [GROUP_DIR, group]
      self._produce_age_file(group_age[group], save_path_trail)
    
    # Creates age data file for overall TDC
    tdc_age = OrderedDict()
    for issue in self.issue_types:
      tdc_age[issue] = self.calc.combine_age_data(group_age, issue)
    save_path_trail = self.base_dir_trail + [GROUP_DIR, TDC_DIR]
    self._produce_age_file(tdc_age, save_path_trail)
    
if (__name__=='__main__'):
  report = JiraGTReport()
  report.produce_report()