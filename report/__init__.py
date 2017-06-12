"""
This package contains modules for running various types of reports.
"""

# Built-in modules
from collections import OrderedDict
from time import time

# User-define modules
from constants import *
from directories import STATE_OF_QUALITY_DIR, PROJECT_DIR, RAW_DATA_DIR, PREV_DATA_DIR, PREV_AGE_DATA_DIR
from db_accessor.jira_gt import JiraGT
from calculator import JiraGTCalculator
from xl_writer import RawDataWriter, ExportDataWriter, TableDataWriter
from utilities import create_dirpath, move_old_files

def timer(func, action='Operation', *args, **kwargs):
  """
  Decorator function that prints the execution time of the func after 
  executing it.
  
  @param func: Function being called and timed.
  @param action: Textual description of the action being performed by the function.
  @param *args: Arbitrary parameters passed into function.
  @param **kwargs: Arbitrary keyword parameters passed into function.
  """
  
  start_time = time()
  results = func(*args, **kwargs)
  print "\n>>>> %s ran in %f seconds. <<<<\n" % (action, time() - start_time)
  return results

class Report(object):
  """
  Generic class responsible for producing the State of Quality reports.
  """
  
  def __init__(self):
    """
    Initializes the parameters responsible for producing reports for a specific
    data source.
    """
    
    # Database object (needs to be set to an actual _DBAccessor subclass)
    self.db = None
    
    # Calculator class (needs to be set to an actual Calculator subclass)
    self.calc = None
    
    # Powerpoint object (needs to be set to an actual Powerpoint subclass)
    self.ppt = None
    
    # Data source name associated with the report
    self.data_source = '<undefined>'
    
    # Base folder trail (as a list) for all files produced by the report
    self.base_dir_trail = [STATE_OF_QUALITY_DIR]
    
    # Data dictionary mapping project group to project list
    self.project_map = { }
    
    # Raw data headers (a list of list of tuples with header name & cell width)
    self.raw_data_headers = [
      [(KEY, 12), (ISSUETYPE, 16), (PRIORITY, 10), ('Current Status', 16), (CREATED, 17)],
      [(KEY, 12), (OLD, 16), (NEW, 16), (TRANS, 17)]
    ]

  def _get_project_map(self):
    """
    Returns the project group to project list mapping for the current report.
    
    Note: This function's return value should be set to self.project_map by a
    sub-class within its constructor.
    
    @return: Data dictionary mapping project groups to lists of associated
    projects.
    """
    
    raise NotImplemented('Needs to be implemented by sub-class.')
  
  def _query_data(self, *args, **kwargs):
    """
    Iteratively queries data from the data source.
    
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    @yield: A tuple of three items, with the first two items being group name
    and project name, and the third item being a data dictionary containing the
    associated queried data for raw data files and metric Excel reports.
    """
    
    raise NotImplemented('Needs to be implemented by sub-class.')
  
  def _prepare_data_for_raw_file(self, data):
    """
    Helper function that converts the given data dictionary into a format
    suitable for passing into the RawDataWriter class.
    
    @param data: Data dictionary mapping issue keys to their corresponding
    fields and field values. It has the following structure:
      <key> -> <field name> -> <field value>
    @return: Data dictionary mapping data types to corresponding lists of lists
    of parameters. It separates the main data from the historical data within
    the original data dictionary. It has the following structure:
      <data type (main data, change history)> -> [[parameters of single key], ]
    """
    
    raw_data = OrderedDict([(MAIN, []), (CHANGE, [])])
    for key, params in data.iteritems():
      curr_params = [key]
      
      # Obtains main data parameters
      for k, v in params.iteritems():
        if (k != HIST): curr_params.append(v)
      raw_data[MAIN].append(curr_params)
      
      # Gets historical data if it exists (for write to separate sheet)
      if (HIST in params):
        for index, old in enumerate(params[HIST][OLD]):
          raw_data[CHANGE].append([key, old, params[HIST][NEW][index], 
                                   params[HIST][TRANS][index]])
    
    return raw_data
  
  def _produce_raw_data_file(self, project, data, *args, **kwargs):
    """
    Produces the Excel raw data file storing the given data.
    
    @param project: Name of project for which raw data files are being produced.
    @param data: Data being stored as raw data.
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    @return: File path of the raw data file produced.
    """
    
    # Save path for raw data
    directories = self.base_dir_trail + [PROJECT_DIR, project, RAW_DATA_DIR]
    save_path = create_dirpath(subdirs=directories)
      
    # Prepares data dictionary for write to the sheets within new Excel file
    raw_data = self._prepare_data_for_raw_file(data)
          
    # Writes data to raw data files
    raw_writer = RawDataWriter(raw_data.keys(), self.raw_data_headers, 
                               '%s Raw Data' % project, save_path)
    return raw_writer.produce_workbook(raw_data)
  
  def _produce_age_file(self, data, save_path_trail, prefix='Average Issue',
                        chart_title='Average Aging (in days)', side_header='Priority',
                        top_header='Status'):
    """
    Produces a single Excel file containing age data for all available issue
    types.
    
    @param data: Data being segmented for the age data file.
    @param save_path_trail: List of directories that form a trail to the save
    location, starting from the base Files folder (which does not need to be
    included in the list).
    @param prefix: Prefix to the given file name.
    @param chart_title: Title of the table within the sheets.
    @param side_header: Header for the side of the table.
    @param top_header: Header for the top of the table.
    @return: File path of the age data file produced.
    """
    
    # Save path for project
    save_path = create_dirpath(subdirs=save_path_trail)
    
    # Moves all old files of the given data type to the old folder
    move_old_files([AGE], save_path_trail + [PREV_AGE_DATA_DIR])
    
    # Sets file name (without data or file extension)
    file_name = '%s %s' % (prefix, AGE)
    
    # Saves age data
    writer = TableDataWriter(file_name, save_path, chart_title=chart_title, 
                        side_header=side_header, top_header=top_header)
    return writer.produce_workbook(data, write_func=float)
  
  def _produce_chart_file(self, project, data, issue_type, data_type, 
                          series_names, prefix=''):
    """
    Produces a single metric Excel file containing chart data.
    
    @param project: Name of project that the file is associated with.
    @param data: Data to be inserted into the metric file.
    @param issue_type: Type of issue represented within metric file (Defect,
    Change Request, Task, etc).
    @param data_type: Type of data being produced by chart (Status, Severity, 
    etc).
    @param series_names: List of tuple containing series names paired with 
    their associated descriptions.
    @param prefix: Prefix to the given file name.
    @return: File path of the metric data file produced.
    """
    
    # Save path for project
    directories = self.base_dir_trail + [PROJECT_DIR, project]
    save_path = create_dirpath(subdirs=directories)
    
    # Moves all old files of the given data type to the old folder
    move_old_files([data_type], directories + [PREV_DATA_DIR])
    
    # Determines file name based on issue type and data type
    file_name = '%ss by %s' % (issue_type, data_type)
    if (prefix): file_name = '%s %s' % (prefix, file_name)
    
    # Determines sheet names and fill map based on data type
    fill_map = { }
    if (data_type == SEV): 
      sheet_names = [TOTAL, OPEN, CLOSED]
      # Backwards because xlsxwriter series colors get reversed based on insertion
      colors = ['blue', 'green', 'yellow', 'orange', 'red']
      fill_map = { series : { 'color' : color } 
                for (series, _), color in zip(series_names, colors) }
    elif (data_type == STATUS): 
      sheet_names = [TOTAL]
    else: 
      sheet_names = []
      
    # Sets up sheet parameters
    sheet_data = OrderedDict([(sheet, { }) for sheet in sheet_names])
    for sheet in sheet_names:
      # Chart parameters
      chart1={'chart_name' : file_name, 'issue_type' : issue_type, 'fill_map' : fill_map}
      chart2={'chart_name' : file_name, 'issue_type' : issue_type, 'fill_map' : fill_map,
            'insertion_row' : 39 + (2 * len(series_names)), 'weeks' : 26}
    
      # Sets chart parameters
      sheet_data[sheet]['chart_params'] = [chart1, chart2]
      
    # Performs exports
    exporter = ExportDataWriter(file_name, save_path, series_names)
    return exporter.produce_workbook(data, sheet_data=sheet_data)
  
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
    
    raise NotImplemented('Needs to be implemented by sub-class.')
  
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
    
    raise NotImplemented('Needs to be implemented by sub-class.')
  
  def _generate_powerpoint(self, *args, **kwargs):
    """
    Uses existing Excel files to produce a Powerpoint report containing all the
    calculated metric data.
    
    @param *args: Arbitrary list arguments for the function.
    @param *kwargs: Arbitrary keyword arguments for the function.
    @return: File path for the Powerpoint file.
    """
    
    self.ppt.generate_presentation()
  
  def produce_report(self):
    """
    Produces the State of Quality report for the associated data source.
    """
    
    # Connects into database
    self.db.connect()
    
    # Prepares aging data data dictionary
    age_data = OrderedDict()
    
    # Queries data for use
    for group, project, data in timer(self._query_data, action='Querying data'):
      if (group not in age_data): age_data[group] = OrderedDict()
      if (project not in age_data[group]): age_data[group][project] = OrderedDict()
      
      # Produces Excel raw data files
      timer(self._produce_raw_data_file, project=project, data=data, 
                            action='Producing raw data files for %s' % project)
      
      # Produces Excel metric data for Powerpoint presentation (and gets age data)
      age_data[group][project] = timer(self._produce_metric_files, project=project, 
        data=data, action='Producing metric files for %s' % project)[2]
                                
    # Disconnects from database
    self.db.disconnect()
                                
    # Produces age data charts for project groups and overall dictionary
    self._produce_group_age_tables(age_data, action='Producing group-level aging files')
    
    # Generates Powerpoint
    timer(self._generate_powerpoint, action='Producing Powerpoint presentation')