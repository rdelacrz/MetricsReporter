"""
Produces report specifically for ClearQuest.
"""

# Built-in modules
from collections import OrderedDict
import xlrd

# User-defined modules
from calculator import DCRCalculator, RRCalculator, SCRCalculator
from constants import *
from directories import RAW_DATA_DIR, PROJECT_DIR, MATRICES_DIR, CQ_DIR, GROUP_DIR, TMS_DIR
from db_accessor import ClearQuest
from powerpoint import ClearQuestPPT
from report import Report
from utilities import create_dirpath
from xl_writer import RawDataWriter

# Set to determine whether project groupings are pulled from Excel or not
USE_EXCEL = True

class ClearQuestReport(Report):
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
    super(ClearQuestReport, self).__init__()
    
    # Sets database to ClearQuest instance
    self.db = ClearQuest()
    
    # Sets calculator to the three ClearQuest instances
    self.calc = OrderedDict([
      (DCR, DCRCalculator), (RR, RRCalculator), (SCR, SCRCalculator)
    ])
    
    # Sets project map
    self.project_map = project_map if (project_map) else self._get_project_map()
    
    # Sets Powerpoint generator
    self.ppt = ClearQuestPPT(self.project_map)
    
    # Data source name associated with the report
    self.data_source = 'ClearQuest'
    
    # Appends to base trail
    self.base_dir_trail += [CQ_DIR]
    
    # Issue types associated with the report
    self.issue_types = OrderedDict([
      (DCR, [ENG_CHANGE, ENG_NOTICE]), (RR, [DEV, PROD]), (SCR, [DEFECT, ENHANCE])
    ])
    
    # Issue type headers
    self.raw_data_headers = OrderedDict([
      (DCR,
        [[(KEY, 12), (SUBMITTER, 20), (REL_TYPE, 16), (STATUS, 16), 
         (LINKED_DOC, 20), (SUBMIT_DATE, 17)],
        [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 17)]]
      ),
      (RR,
        [[(KEY, 12), (HEADLINE, 25), (REL_TYPE, 16), (STATUS, 16),
         (PROPERTY, 17), (SUBMIT_DATE, 17), (CLOSED_DATE, 17), 
         (EST_FIELD_DATE, 20), (ACT_FIELD_DATE, 20)],
        [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 17)]]
      ),
      (SCR,                             
        [[(KEY, 12), (HEADLINE, 25), (ISSUETYPE, 16), (PRIORITY, 10), 
         (STATUS, 16), (LINKS, 20), (PROPERTY, 17), (SUBMIT_DATE, 17), 
         (CLOSED_DATE, 17), (EST_FIX_TIME, 20), (ACT_FIX_TIME, 20)],
        [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 17)]]
      )
    ])
    
  def _read_excel_project_map(self):
    """
    Reads the group to project mappings from the JIRA Projects Excel file.
    
    @return: Data dictionary mapping project group names to a list of tuples
    with project keys and associated full project names. It looks like:
      <project group> -> [(project key, full project name),]
    """
    
    # Initializes group map
    group_map = { TMS_DIR : [] }
    
    # File path of JIRA projects file
    filepath = '%s\\TMS Projects.xlsx' % create_dirpath(subdirs=[MATRICES_DIR])
    
    # Workbook for file
    wb = xlrd.open_workbook(filepath)
    
    # Sheet with project information
    sheet = wb.sheet_by_index(0)
    
    # Traverses through each row for projects
    for row in range(sheet.nrows)[1:]:
      name = str(sheet.cell(row, 0).value).strip()
      in_use = str(sheet.cell(row, 2).value).strip()
      if (in_use == 'Yes'):
        group_map[TMS_DIR].append((name, name))
    
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
        # Gets DCR, RR, and SCR data
        data = OrderedDict([
          (DCR, self.db.get_dcr_data(proj)),
          (RR, self.db.get_rr_data(proj)),
          (SCR, self.db.get_scr_data(proj)),
        ])
          
        yield group, proj, data
        
  def _produce_raw_data_file(self, project, data, *args, **kwargs):
    """
    Produces the Excel raw data files storing the given data.
    
    @param project: Name of project for which raw data files are being produced.
    @param data: Data being stored as raw data.
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    @return: File paths of the raw data files produced.
    """
    
    # Save path for raw data
    directories = self.base_dir_trail + [PROJECT_DIR, project, RAW_DATA_DIR]
    save_path = create_dirpath(subdirs=directories)
      
    file_paths = []
    for metric_type, metric_data in data.iteritems():
      # Prepares data dictionary for write to the sheets within new Excel file
      raw_data = self._prepare_data_for_raw_file(metric_data)
            
      # Writes data to raw data files
      raw_writer = RawDataWriter(raw_data.keys(), self.raw_data_headers[metric_type], 
                        '%s %s Raw Data' % (project, metric_type), save_path)
      file_paths.append(raw_writer.produce_workbook(raw_data))
    return file_paths
    
  def _produce_metric_files(self, project, data, *args, **kwargs):
    """
    Produces all the Excel metric files to be used in the final Powerpoint 
    report.
    
    @param project: Name of project for which metric files are being produced.
    @param data: Data from which calculated metrics are derived.
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    @return: Tuples of three data dictionaries for severity data, status data, 
    and age data.
    """
    
    # Initializes overall metric data dictionaries
    all_sev_data = OrderedDict()
    all_status_data = OrderedDict()
    all_age_data = OrderedDict()
    
    # Iterates through each data type
    for metric_type, metric_data in data.iteritems():
      # Initializes segmented data dictionary
      segmented_data = OrderedDict([(x, OrderedDict()) 
                                    for x in self.issue_types[metric_type]])
      
      # Segments data by issue type
      for key, params in metric_data.iteritems():
        issue_type = params.get(ISSUETYPE, params.get(REL_TYPE))
        segmented_data[issue_type][key] = params
        
      # Performs calculations for each issue type
      sev_data = OrderedDict()
      status_data = OrderedDict()
      age_data = OrderedDict()
      for issue_type in self.issue_types[metric_type]:
        calc = self.calc[metric_type](project, issue_type)
        sev, status, age = calc.get_metrics(segmented_data[issue_type])
        
        # Inserts data into data dictionaries
        sev_data[issue_type] = sev
        status_data[issue_type] = status
        age_data[issue_type] = age
        
        # Produces severity file (if priorities exist)
        if (calc.priority_list):
          series_names = [(p, p) for p in calc.priority_list]
          self._produce_chart_file(project, sev, issue_type, SEV, series_names, 
                                   prefix=metric_type)
        
        # Produces status file
        series_names = calc.get_status_desc()
        self._produce_chart_file(project, status, issue_type, STATUS, series_names, 
                                 prefix=metric_type)
        
      # Produces age file
      self._produce_age_file(age_data, self.base_dir_trail + [PROJECT_DIR, project],
                             prefix='Average %s' % metric_type)
      
      # Maps data to data dictionaries
      all_sev_data[metric_type] = sev_data
      all_status_data[metric_type] = status_data
      all_age_data[metric_type] = age_data
      
    return all_sev_data, all_status_data, all_age_data
  
  def _produce_group_age_tables(self, data, *args, **kwargs):
    """
    Combines data dictionary age data of every group contained inside, and 
    outputs Average Age data file for all of them.
    
    @param data: Multi-level data dictionary mapping groups to projects to age
    data dictionaries. It has the following structure:
      <project group> -> <project> -> <metric type> -> <issue type> -> <priority> 
        -> <status group> -> <AverageAge object>
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    """
    
    # Regroups data for use with combine_age_data() function
    regrouped_data = OrderedDict()
    for group, project_map in data.iteritems():
      regrouped_data[group] = OrderedDict()
      for project, metric_map in project_map.iteritems():
        for metric_type, age_data in metric_map.iteritems():
          if (metric_type not in regrouped_data[group]):
            regrouped_data[group][metric_type] = OrderedDict()
          regrouped_data[group][metric_type][project] = age_data
          
    # Initializes project group level aging data
    group_age = OrderedDict([(group, OrderedDict([
      (metric_type, OrderedDict()) for metric_type in self.issue_types.keys()
    ])) for group in data])
    
    # Creates age files for all project groups
    for group, metric_map in regrouped_data.iteritems():
      for metric_type, issue_list in self.issue_types.iteritems():
        for issue in issue_list:
          group_age[group][metric_type][issue] = \
            self.calc[metric_type].combine_age_data(metric_map[metric_type], issue)
        
        # Creates age data file for project group and metric type
        age_data = group_age[group][metric_type]
        save_path_trail = self.base_dir_trail + [GROUP_DIR, group]
        self._produce_age_file(age_data, save_path_trail, prefix='Average ' + metric_type)
        
if (__name__=='__main__'):
  report = ClearQuestReport()
  report.produce_report()