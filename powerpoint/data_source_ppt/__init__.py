"""
This package contains multiple modules for producing Powerpoint presentations
for each available data source.
"""

# Built-in modules
from collections import OrderedDict
import os
import re

# User-defined modules
from constants import BLOCKER, CRITICAL, MAJOR, MINOR, TRIVIAL, AGE, SEV, STATUS, TOTAL
from powerpoint import Powerpoint
from utilities import get_str_date

# Sets aging slide note (added below the aging table)
AGING_NOTES = """* The Overall column for statuses corresponds to the average \
overall age of all the issues. The age of each individual issue is equal to the \
amount of time that has elapsed between the issue's creation date and closure date. \
If an issue is not closed, the current date is used instead of closure date.
* The Overall row for priorities corresponds to all issues within a given status \
group, regardless of what priority it has."""

class DataSourcePPT(Powerpoint):
  """
  Generic class for producing Powerpoint presentations for each data source.
  """
  
  def __init__(self, data_source='<undefined>'):
    """
    Initializes parameters for creating the Powerpoint.
    
    @param data_source: Name of the data source (defaults to undefined).
    """
    
    super(DataSourcePPT, self).__init__()
    
    # Sets data source name for logging
    self.data_source = data_source
    
    # Sets title for presentation (displayed on title slide)
    self.title = '"State of Quality" \nAssessment'
    
    # Sets base file name
    self.file_name = 'State of Quality Assessment'
    
    # Sets issue type list (empty by default)
    self.issue_types = []
    
    # Sets issue type definitions (empty by default)
    self.issue_def = { }
    
    # Sets project and project group file paths (None by default)
    self.project_path = None
    self.project_group_path = None
    
  _row_col_re = re.compile('# of rows: (\d+?), # of columns: (\d+)')
  def has_data(self, sheet, data_type=AGE):
    """
    Checks the given Excel sheet if it has any data. The way it checks the
    sheet depends on the type of data being read within the sheet.
    
    @param sheet: Excel sheet being checked for data.
    @param data_type: The type of data contained within the Excel sheet.
    @return: True if data exists in sheet, False otherwise.
    """
    
    # Reading an Age Data table
    if (data_type == AGE):
      text_obj = self._row_col_re.search(sheet.Cells(1,1).Value)
      rows = int(text_obj.group(1))
      cols = int(text_obj.group(2))
      data = int(sheet.Cells(rows + 2, cols + 1).Value)    # Overall
    # Reading a status or severity table
    elif (data_type in [SEV, STATUS]):
      data = int(sheet.Cells(1, 2).Value)
    else:
      raise ValueError('Data type passed not recognized.')
    
    return (data > 0)     # True if data is greater than zero
    
  def add_priority_table_slide(self, data_source, title='Things to Know (continued)',
                         log='Priority table slide added.'):
    """
    Adds a slide discussing each priority within a table.
    
    @param data_source: Name of the associated data source (front end name).
    @param title: Title of priority slide.
    @param log: Log information to be printed on the screen.
    @return: Slide object.
    """
    
    # Add initial slide
    slide = self.add_slide(self.blank_index, title=title, log=log)
    
    # Prepares priority table
    table_map = OrderedDict([
      (('Priority', 140), ['Blocker', 'Critical', 'Major', 'Minor', 'Trivial']),
      (('Definition', 400), [
        """Highest priority - takes precedence over all others - creating \
revenue loss for our clients' operations; no workaround is available""",
        """This is causing a problem and requires urgent attention - may be \
impacting project schedule/project delivery; no workaround is available""",
        """This has a significant impact - a workaround may be available; but \
total solution needs to be in place before delivery""",
        """This has a relatively minor impact""",
        """Lowest priority"""
      ]),
      (('Closure Goal', 140), ['Within 2 days', 'Within 2 days', 'TBD', 'TBD', 'TBD'])
    ])
    
    # Adds priority table to slide
    table = self.add_ppt_table(slide, table_map, merges=[(7,1,7,3)],
                       NumRows=7, NumColumns=3, Left=20, Top=80, Height=100)
    
    # Adds additional final row text to table
    text = 'Guideline:  priority indicates its importance relative to the' + \
                ' other open %s tickets for that project' % data_source
    table.Cell(7, 1).Shape.TextFrame.TextRange.Text = text
    word_count = len(data_source.split())
    table.Cell(7, 1).Shape.TextFrame.TextRange.Words(7, 9 + word_count).Font.Underline = True
    table.Cell(7, 1).Shape.TextFrame.TextRange.Words(7, 9 + word_count).Font.Italic = True
    
    return slide
  
  def add_status_table_slide(self, issue_type, status_map, title='Things to Know (continued)',
                       log='Priority table slide added.'):
    """
    Adds a slide discussing the various status groups for the given issue type.
    
    @param issue_type: Type of issue being discussed in the slide.
    @param status_map: Ordered data dictionary mapping each status group to the 
    list of statuses associated with that given group.
    @param title: Title of status slide.
    @param log: Log information to be printed on the screen.
    @return: Slide object.
    """
    
    # Add initial slide
    slide = self.add_slide(self.blank_index, title=title, log=log)
    
    # Adds issue type as secondary header
    self.add_secondary_header(slide, issue_type, top=60)
    
    # Sets headers
    group_header = 'Status Group'
    statuses_header = 'Statuses Included'
    
    # Prepares status table map
    table_map = OrderedDict([((group_header, 190), []), ((statuses_header, 490), [])])
    for group, desc in status_map.iteritems():
      table_map[(group_header, 190)].append(group)
      table_map[(statuses_header, 490)].append(desc)
      
    # Adds status table to slide
    self.add_ppt_table(slide, table_map, NumRows=len(status_map) + 1, 
                       NumColumns=2, Left=20, Top=120, Height=100)
    
    return slide
  
  def add_project_group_slide(self, project_group, layout_index=22, 
                              log='Project group slide added.'):
    """
    Adds an introductory project group slide.
    
    @param project_group: Name of the project group being added.
    @param layout_index: Layout index to be used for project group slide.
    @param log: Log information to be printed on the screen.
    @return: Slide object.
    """
    
    return self.add_slide(layout_index, title=project_group, logo=False, 
                          footer=None, log=log)
    
  def get_source_str(self, date):
    """
    Determine the source data string to be printed at the bottom of slides,
    based on the current database and given source data date.
    
    @param date: Datetime object representing the date of associated source 
    data (when data was last extracted).
    """
    
    # Gets source string
    date_str = date.strftime('%m/%d/%Y')
    produced =' Produced by: Quality Assurance for Service Delivery'
    return 'Source: %s, %s\n%s' % (self.data_source, date_str, produced)
  
  def _add_issue_type_def(self, slide, issue_type):
    """
    Adds the definition of the given issue type to the top right of the slide,
    if a definition exists.
    
    @param slide: The slide to which the definition is being added.
    @param issue_type: The issue type whose definition is being placed onto the
    slide.
    """
    
    # Only adds definition if issue type exists in issue_def data dictionary
    if (issue_type in self.issue_def):
      textframe = self.add_textframe(slide, self.issue_def[issue_type], 11, 0, False,
                         Orientation=1, Left=420, Top=25, Width=280, Height=50)
      textframe.TextRange.Words(2,1).Font.Color.RGB = self.colornum   # Partially colored
      
      # Resizes title box accordingly
      slide.Shapes.Title.Width = 400
      
  def _add_source_info(self, slide, info):
    """
    Adds source information to the bottom of the slide, if any is given.
    
    @param slide: The slide to which the source info is being added.
    @param info: The source information being added to the slide.
    """
    
    if (info): self.add_textframe(slide, info, 14, 0, False, Orientation=1, 
                                  Left=210, Top=485, Width=350, Height=50)
  
  def add_aging_table_slide(self, sheet, issue_type, main_header, sub_header=None,
        left=30, note=AGING_NOTES, source='', log='Aging table slide added.'):
    """
    Adds a slide for average aging, using the table inside the given Excel
    sheet.
    
    @param sheet: Excel sheet containing average aging data for a given project
    or project group.
    @param issue_type: Issue type that the aging data is representing.
    @param main_header: Main header to add to the top of the slide.
    @param sub_header: Sub header to add below the main header.
    @param left: How far from the left side of the slide should the table be
    inserted.
    @param note: Additional note to be added below the chart.
    @param source: Source information to be added to the bottom of the slide.
    @param log: Log information to be printed on the screen.
    @return: Slide object.
    """
    
    # Add initial slide
    slide = self.add_slide(self.blank_index, title=main_header, log=log)
    
    # Sets sub header
    if (sub_header): self.add_secondary_header(slide, sub_header)
    
    # Gets the number of rows and columns of table (printed on top left)
    row_col_info = sheet.Cells(1,1).Value
    rows = int(self._row_col_re.search(row_col_info).group(1))
    cols = int(self._row_col_re.search(row_col_info).group(2))
    
    # Pastes the table from the Excel sheet into the slide
    table = self.paste_excel_range(sheet, slide, 1, 2, cols, 1 + rows, 
                      width=self.slide_width - (2 * left), left=left, top=150)
    
    # Adds the aging note below the inserted chart
    self.add_textframe(slide, note, 12, 0, False, Orientation=1, Left=20, 
                       Top=table.Height + 155, Width=680, Height=50)
    
    # Adds the issue type definition to the top right, if any
    self._add_issue_type_def(slide, issue_type)
    
    # Adds the source information, if any
    self._add_source_info(slide, source)
    
    return slide
  
  def add_stacked_chart_slide(self, sheet, chart_name, main_header, sub_header, 
            issue_type, left=30, source='', log='Stacked chart slide added.'):
    """
    Adds a slide for status or priority trends, using the stacked chart inside
    the given Excel sheet.
    
    @param sheet: Excel sheet containing average aging data for a given project
    or project group.
    @param chart_name: Name of the chart within the sheet being added.
    @param main_header: Main header to add to the top of the slide.
    @param sub_header: Sub header to add below the main header.
    @param issue_type: Issue type that the aging data is representing.
    @param left: How far from the left side of the slide should the table be
    inserted.
    @param source: Source information to be added to the bottom of the slide.
    @param log: Log information to be printed on the screen.
    @return: Slide object.
    """
    
    # Add initial slide
    slide = self.add_slide(self.blank_index, title=main_header, log=log)
    
    # Sets sub header
    self.add_secondary_header(slide, sub_header)
    
    # Pastes the chart within the Excel sheet into the Powerpoint slide
    self.paste_excel_chart(sheet, chart_name, slide)
    
    # Adds the issue type definition to the top right, if any
    self._add_issue_type_def(slide, issue_type)
    
    # Adds the source information, if any
    self._add_source_info(slide, source)
    
    return slide
  
  def add_individual_priority_slide(self, priority, log='Priority elaboration slide added.'):
    """
    Adds slide about the given priority.
    
    @param priority: Priority being described.
    @return: Slide object.
    """
    
    # Sets definition, detail, and example text based on given priority
    if (priority == BLOCKER):                                            # BLOCKER
      definition = 'Definition: Highest priority. Indicates that this issue takes precedence over all others.'
      detail = """A part of the software isn't functioning as intended and results \
in an inability to perform a critical business process.  Time sensitive items, there is no \
workaround available and the problem is immediate and blocking progress on a project.

A description of "why" the impact of this issue causes a "block" situation for this project. \
The explanation is required only when priority = Blocker."""
      example = ["Missing images about to expire",
                 "Next step in workflow cannot be completed until this issue is resolved",
                 "Issues with limited response time per contract",
                 "System or major components un-usage by multiple users or clients"]
      left = 135
      top = 40

    elif (priority == CRITICAL):                                         # CRITICAL
      definition = 'Definition: Indicates that this issue is causing a problem and requires urgent attention.'
      detail = """A part of the software isn't functioning as intended and results \
in an inability to perform a critical business process. Time sensitive items need addressing, \
there is no workaround available but the problem is not immediately blocking progress on \
a project."""
      example = ["Images missing from download", "Missing deployment logs", "Impact printing",
                 "DMV issues", "Large number of events cannot be processed",
                 "Credit card payment issues",
                 "Set-up new clients/databases/locations/scripts (if timeline is critical)"]
      left = 130
      top = 40
        
    elif (priority == MAJOR):                                             # MAJOR
      definition = 'Definition: Indicates that this issue has a significant impact.'
      detail = """A part of the software isn't functioning as intended and results \
in an inability to perform a business process. Items are not time sensitive. There is a \
workaround available and/or the problem is not immediately blocking progress on a project."""
      example = ["Issues with financial reporting", "Issues/items need to be addressed",
                 "Set-up new clients/databases/locations/scripts (if timeline is not critical)"]
      left = 110
      top = 40

    elif (priority == MINOR):                                             # MAJOR
      definition = 'Definition: Indicates that this issue has a relatively minor impact.'
      detail = """A part of the software isn't functioning as intended but does not \
result in an inability to perform a business process. Items that would be great if \
changed/addressed, there is a workaround available, and the problem is not immediately \
blocking progress on a project."""
      example = ['New functionality/report that is more user friendly']
      left = 110
      top = 35
        
    elif (priority == TRIVIAL):                                             # MAJOR
      definition = 'Definition: Lowest priority.'
      detail = """Cosmetic changes to be approved with budget if time permitting. \
To be reviewed/accepted by change control board and roll out with future software revisions. \
The software is functioning as intended but the UI is not is correct (i.e. too long words, \
wrong colors, not aligned, etc.)."""
      example = ["Add a period at the end of every sentence", 
                 "Change word choice, color, move a button .5 inch to the left"]
      left = 115
      top = 30
        
    else:
      raise ValueError('Unrecognizable priority passed.')
    
    # Creates slide with example bullet points
    offset = 50 if priority == 'Blocker' else 0
    slide = self.add_bullets_slide(priority, example, left=10, top=175 + offset, width=700, 
                                   height=50, text_size=14, text_color=self.colornum, log=log)
    self.add_textframe(slide, 'Examples:', 14, self.colornum, False, 
                       Orientation=1, Left=10, Top=160 + offset, Width=700, Height=10)
    
    # Adds priority symbol
    isopath = self.template_path + '\\priority_%s.png' % priority.lower()
    if (not os.path.isfile(isopath)):
      raise Exception('Missing template file')                    # Means priority file is missing
    slide.Shapes.AddPicture(isopath, 0, -1, left, top, 25, 25)
    
    # Adds priority definition
    textframe = self.add_textframe(slide, definition, 14, self.colornum, False, 
                       Orientation=1, Left=10, Top=70, Width=700, Height=50)
    textframe.TextRange.Words(2, 15).Font.Bold = True
    textframe.TextRange.Words(2, 15).Font.Italic = True
    
    # Adds priority details
    self.add_textframe(slide, detail, 14, self.colornum, False, 
                       Orientation=1, Left=10, Top=100, Width=700, Height=50)
    
    return slide
  
  def _get_age_files(self, dir_path, keyword=None):
    """
    Gets a list of all the aging files within the given directory path.
    
    @param dir_path: Path of directory being searched for Age Data files.
    @param keyword: Additional keyword to search for in file name, if any.
    @return: List of file paths for Age Data.
    """
    
    age_files = []
    
    # Iterates through directory for aging files
    for file_name in os.listdir(dir_path):
      file_path = os.path.join(dir_path, file_name)
      if (AGE in file_name and os.path.isfile(file_path) 
          and file_name.endswith('xlsx') and not file_name.startswith('~$')):
        if (not keyword or keyword in file_name):
          age_files.append(file_path)
        
    return age_files
  
  def add_project_group_aging_slides(self, group_name, header_map, note=AGING_NOTES,
                                     write_group=True):
    """
    Adds all the aging slides for the aging files within the project group path.
    
    @param group_name: Name of project group associated with aging files.
    @param header_map: Data dictionary mapping issue type to headers. If an
    issue type is not a key in the header map, its associated header will
    default to Overall "State of Quality": <issue_type>.
    @param note: Note to write below the aging table.
    @param write_group: Determines whether group name should be written
    onto the slide's sub-header or omitted.
    """
    
    # Navigates through directory for age data files
    group_path = os.path.join(self.project_group_path, group_name)
    age_paths = self._get_age_files(group_path)
      
    # Opens aging data file, if it exists
    for age_path in age_paths:
      wb = self.excel.Workbooks.Open(age_path)
      
      # Gets issue types
      issue_types = []
      
      # Either a list or a data dictionary of lists
      if (isinstance(self.issue_types, list)): 
        issue_types = self.issue_types
      else:
        for metric_type, issue_type_list in self.issue_types.iteritems():
          if (metric_type in age_path):
            issue_types = issue_type_list
            break
      
      # Iterates through the issue type tabs of the age data file
      for issue_type in issue_types:
        sheet = wb.Sheets(issue_type)
        
        # Get headers
        header = header_map.get(issue_type, 'Overall "State of Quality": ' + issue_type)
        
        # Sets source string
        date = get_str_date(age_path, regex='.+(\d{4})-(\d{2})-(\d{2})\.xlsx')
        source = self.get_source_str(date)
        
        # Adds aging table
        log = '%s aging table slide added' % group_name
        if (not write_group):
          self.add_aging_table_slide(sheet, issue_type, header, sub_header=None, 
                                   note=note, source=source, log=log)
        else:
          self.add_aging_table_slide(sheet, issue_type, header, sub_header=group_name, 
                                   note=note, source=source, log=log)
        
      # Closes workbook
      wb.Saved = True; wb.Close()
        
  def add_project_aging_slides(self, project_key, project_name, issue_type, 
                               note=AGING_NOTES, header=None, keyword=None):
    """
    Adds all the aging slides for the project aging files within the project path.
    
    @param project_key: Key of project associated with aging files.
    @param project_name: Name of project associated with aging files.
    @param issue_type: Issue type associated with the target aging table.
    @param note: Note to write below the aging table.
    @param header: Header to use for the file, if not the issue type.
    @param keyword: Keyword that should be in the aging file, if any.
    """
    
    # Navigates through directory for age data files
    group_path = os.path.join(self.project_path, project_key)
    age_paths = self._get_age_files(group_path, keyword=keyword)
    
    # Opens aging data file, if it exists
    for age_path in age_paths:
      wb = self.excel.Workbooks.Open(age_path)
      sheet = wb.Sheets(issue_type)
        
      # Skips sheet if it has no data
      if (self.has_data(sheet)):
        # Sets header
        header = '%s Average Age Data' % issue_type if (not header) else header
        
        # Sets source string
        date = get_str_date(age_path, regex='.+(\d{4})-(\d{2})-(\d{2})\.xlsx')
        source = self.get_source_str(date)
        
        # Adds aging table
        self.add_aging_table_slide(sheet, issue_type, header, 
                            sub_header=project_name, note=note, source=source, 
                            log='%s aging table slide added' % project_key)
      
      # Closes workbook
      wb.Saved = True; wb.Close()
      
  _issue_type_re = re.compile('(.+?)s by .+?\.xlsx')
  def _get_chart_files(self, dir_path, keywords, issue_type=None):
    """
    Gets an Ordered data dictionary of all the chart files (Severity, Status, etc) 
    of the given issue type within the given directory path, with each keyword 
    within the keywords list mapped to the associated file paths with the keyword.
    
    @param dir_path: Path of directory being searched for chart files.
    @param keywords: List of keywords for file paths to have. If a file path 
    contains any of the keywords, it will be mapped to that keyword.
    @param issue_type: Issue type associated with the target files.
    @return: Data dictionary mapping keywords to associated file paths.
    """
    
    # Initializes paths dictionary
    paths = OrderedDict([(keyword, []) for keyword in keywords])
    
    # Iterates through directory for aging files
    for file_name in os.listdir(dir_path):
      file_path = os.path.join(dir_path, file_name)
      if (os.path.isfile(file_path) and file_name.endswith('xlsx') and not 
          file_name.startswith('~$')):
        # Pulls issue type from file
        file_type_pattern = self._issue_type_re.search(file_name)
        if (file_type_pattern):
          file_type = file_type_pattern.group(1)
          
          # Iterates through keywords to see if file path has any of them
          for keyword in keywords:
            if (keyword in file_path and (not issue_type or file_type==issue_type)):
              paths[keyword].append(file_path)
        
    return paths
      
  def add_project_chart_slides(self, project_key, project_name, issue_type, 
                               data_type, header_map=None):
    """
    Adds all the chart slides for the project charts files within the project 
    path.
    
    @param project_key: Key of project associated with aging files.
    @param project_name: Name of project associated with chart files.
    @param issue_type: Issue type associated with the target chart table.
    @param data_type: Type of data within the target chart.
    @param header_map: Multi-level data dictionary mapping issue type to data 
    type to headers. If header map is not passed, or issue type doesn't exist 
    in the map, a specific default will be used based on data type.
    """
    
    # Navigates through directory for chart data files
    project_path = os.path.join(self.project_path, project_key)
    chart_paths = self._get_chart_files(project_path, [data_type], 
                                        issue_type=issue_type)[data_type]
    
    # Iterates through chart files and adds a slide for each of them
    for chart_path in chart_paths:
      wb = self.excel.Workbooks.Open(chart_path)
      sheet = wb.Sheets(TOTAL)    # Gets total data
      
      # Skips sheet if it has no data
      if (self.has_data(sheet, data_type)):
        # Sets header (uses default if no header map is passed)
        if (data_type == SEV):
          header = '%s Priority Data (Cumulative View)' % issue_type
        elif (data_type == STATUS):
          header = '%s Workflow Activity (Cumulative View)' % issue_type
        else:
          header = '%s Data'        # Default
        if (header_map and issue_type in header_map):
          header = header_map[issue_type].get(data_type, header)
        
        # Sets source string
        date = get_str_date(chart_path, regex='.+(\d{4})-(\d{2})-(\d{2})\.xlsx')
        source = self.get_source_str(date)
        
        # Sets chart name (uses past 6 months charts)
        chart_name = '%s by %s (past 6 months)'
        
        # Adds 6 month chart
        self.add_stacked_chart_slide(sheet, chart_name, header, project_name, 
                              issue_type, source=source,
                              log='%s stacked chart slide added' % project_key)
      
      # Closes workbook
      wb.Saved = True; wb.Close()
  
  def generate_presentation(self, date):
    """
    Generates the full Powerpoint presentation.
    
    @param date: Datetime object representing the date of the presentation, 
    defaulting to the current date.
    """
    
    raise NotImplementedError('Must be implemented by sub-class.')
  
# Imports sub-modules and places them on current level
from gt_jira_ppt import GTJiraPPT
from clearquest_ppt import ClearQuestPPT