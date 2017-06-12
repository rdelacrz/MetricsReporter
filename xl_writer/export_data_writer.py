"""
This modules generates item metrics in the form of charts inside Excel
workbooks.
"""

# Built-in modules
from collections import OrderedDict
from datetime import timedelta

# User-defined modules
from constants import TOTAL
from utilities import create_dirpath
from xl_writer import XLWriter

class ExportDataWriter(XLWriter):
  """
  This class generates Excel metrics charts for a given type of data.
  """
  
  def __init__(self, filename='Export Data', save_path=create_dirpath(), series_names=[]):
    """
    Initializes the Excel file containing raw data, which will be saved at the 
    given path with the given file name.
    
    @param filename: Name of the file (without file extension).
    @param save_path: File path where Excel sheet will be saved at.
    @param series_names: List of tuples that will be used for the series on the
    charts. Each tuple contains the series name and a second string describing
    the series: (series name, series description).
    """
    
    super(ExportDataWriter, self).__init__(filename, save_path)
    self.data_type = 'Export'
    
    # Cell formats
    self.itemnameformat = self.wb.add_format({
      'bold':True, 'font_color':'black', 'font_size':10, 'font':'Arial', 'bg_color':'#BFBFBF'
    })
    self.itemrowformat = self.wb.add_format({
      'font_color':'black', 'font_size':10, 'font':'Arial', 'bg_color':'#BFBFBF'
    })
    self.weeknumformat = self.wb.add_format({
      'font_color':'black', 'font_size':10, 'font':'Arial', 'bg_color':'#E26B0A'
    })
    self.totalnameformat = self.wb.add_format({
      'bold':True, 'font_color':'black', 'font_size':10, 'font':'Arial', 'bg_color':'#92D050'
    })
    self.totalrowformat = self.wb.add_format({
      'font_color':'black', 'font_size':10, 'font':'Arial','bg_color':'#92D050'
    })
    self.dateformat = self.wb.add_format({
      'font_color':'black', 'font_size':10, 'font':'Arial', 'num_format':'m/d/yyyy', 
      'border':1, 'rotation':60
    })
    self.regformat = self.wb.add_format({
      'font_color':'black', 'font_size':11, 'font':'Arial'
    })
    
    # Initializes the series names
    self.series_names = series_names
    
    # Initializes dictionary of sheets
    self.sheets = OrderedDict()
    
    # Sets write rows
    self.date_row = 3
    self.week_row = 4
    self.top_write_row = 6
    
  def _initialize_sheet(self, sheet_name):
    """
    Initializes sheet so that it is write-ready.
    
    @param sheet_name: Name of the sheet being added.
    @return: The sheet object that was initialized.
    """
    
    # Creates the sheet
    write_name = sheet_name[:31] if (len(sheet_name) > 31) else sheet_name
    self.sheets[sheet_name] = self.wb.add_worksheet(write_name)
    
    # Widens the first column
    self.sheets[sheet_name].set_column('A:A', 19)
    
    # Sets the date row format
    self.sheets[sheet_name].set_row(self.date_row, cell_format=self.dateformat)
    
    # Sets the week number row format
    self.sheets[sheet_name].set_row(self.week_row, cell_format=self.weeknumformat)
    
    # Sets the series header and row format
    row = self.top_write_row
    for series_name, _ in self.series_names:
      self.sheets[sheet_name].set_row(row, cell_format=self.itemrowformat)
      self.sheets[sheet_name].write(row, 0, series_name, self.itemnameformat)
      row += 2
    
    # Sets the total header and row format
    self.sheets[sheet_name].write(row, 0, TOTAL, self.totalnameformat)
    self.sheets[sheet_name].set_row(row, cell_format=self.totalrowformat)
    
    return self.sheets[sheet_name]
  
  def _create_chart(self, sheet, final_col, chart_name, issue_type='Issue', 
                    insertion_row=None, weeks=None, fill_map=None, min_y=None):
    """
    Creates chart on the current sheet based on the given parameters.
    
    @param sheet: The sheet containing the data that the charts will be based on.
    @param final_col: The final column of the data.
    @param chart_name: The name of the chart to be displayed on the top.
    @param issue_type: The type of issue being measured by the chart.
    @param insertion_row: The row at which the chart is inserted.
    @param weeks: Can be an integer or a dictionary mapping sheet names to 
    integers. The integer value represents the maximum number of weeks to be 
    covered by the chart. If no number of weeks is given, the full range of the
    chart is covered.
    @param fill_map: Data dictionary mapping series names to their 
    corresponding fill parameters.
    @param min_y: Can be an integer or a dictionary mapping sheet names to 
    integers. The integer value represents the minimum value on the y-axis of
    the chart.
    """
    
    # Initializes x-axis parameters
    x_axis = {
      'name': 'Dates',
      'name_font': {'size': 12, 'bold': True, 'font': 'Arial'},
      'num_font':  {'italic': True, 'font': 'Arial'},
      'date_axis': True, 'num_format': 'mm/dd/yyyy'
    }
    
    # Updates chart name if weeks is given
    if (weeks):
      if (weeks == 26):
        chart_name += ' (past 6 months)'
      elif (weeks == 9):
        chart_name += ' (past 2 months)'
      else:
        chart_name += ' (past %d weeks)' % weeks
        
      # Separate code changing x-axis information based on week range
      if (weeks <= 26):
        x_axis.update({
          'major_unit': 7, 'major_unit_type': 'days', 
          'minor_unit': 1, 'minor_unit_type': 'days'
        })
    
    # Creates the chart
    chart = self.wb.add_chart({'type':'area', 'subtype':'stacked', 'name':chart_name})
    chart.set_size({'width': 1200, 'height': 600})
    chart.set_title({
      'name': chart_name,
      'name_font': {'name': 'Arial'}
    })
    chart.set_y_axis({
      'name': 'Number of %ss' % issue_type,
      'name_font': {'size': 12, 'bold': True, 'font': 'Arial'},
      'min' : min_y
    })
    chart.set_x_axis(x_axis)
    
    # Determines starting column of range
    first_col = 1 if (not weeks or final_col <= weeks) else (final_col - weeks)
    
    # Sets insertion row for chart, if it has not been set
    if (not insertion_row):
      insertion_row = 8 + (2 * len(self.series_names))
    
    # Iterates through each item being counted
    row = 4 + (2 * len(self.series_names))
    for series_name, _ in self.series_names:
        # Adds series to the chart based on the current row of data
        chart.add_series({
            'name' : "='%s'!$A$%d" % (sheet.name, row+1),
            'categories': ["'%s'" % sheet.name, 3, first_col, 3, final_col],
            'values': ["'%s'" % sheet.name, row, first_col, row, final_col],
            'fill': None if (not fill_map) else fill_map[series_name],
        })
        
        row -= 2
        
    # Insert chart into the sheet below the data
    sheet.insert_chart(insertion_row, 1, chart)
  
  def _finalize_sheet(self, sheet, final_col):
    """
    Performs final activities for the current sheet, after all the data has
    been written onto the sheet.
    
    @param sheet: Sheet with the data written.
    @param final_col: The final column containing count data.
    """
    
    # Adds week number to the top of the sheet
    sheet.write('A1', 'Number of Weeks:', self.regformat)
    sheet.write('B1', final_col, self.regformat)
    
    # Resizes the widths of the columns
    sheet.set_column(1, final_col, 4)
    
    # Obtains the row of the total count and writes Total on the right
    row = self.top_write_row + (2 * len(self.series_names))
    sheet.write(row, final_col + 2, TOTAL, self.totalnameformat)
    
    # Writes the description of each item on the right of the counts
    for _, desc in reversed(self.series_names):
      row -= 2
      sheet.write(row, final_col + 2, desc, self.itemnameformat)
      
  def _add_sheet(self, sheet_name, data, omit_last=False, date_shift=None, chart_params=[]):
    """
    Populates sheet with the given data.
    
    @param sheet_name: Name of the sheet being added - the higher level data type.
    @param data: Data dictionary whose data will populate the sheet. It has the
    following structure:
      <date> -> <sheet name> -> <series name> -> count
    @param omit_last: Denotes whether most recent week's data should not be 
    added to the sheet.
    @param date_shift: Denotes the number of days that the displayed days
    should be shifted, if any.
    @param chart_params: A list of data dictionaries, with each one containing
    various parameters for each chart to be generated within the chart.
    """
    
    # Initializes sheet for use
    sheet = self._initialize_sheet(sheet_name)
    
    # Initializes writing column number
    col = 1
    
    # Determines when an issue is finally found during the date iteration
    data_found = False
    
    # Creates the list of dates to iterate through
    dates = data.keys()[:-1] if (omit_last) else data.keys()
    
    # Writes the counts on the given sheet for every date
    for date in dates:
      # Skips any dates with no data
      if (data_found or data[date][sheet_name][TOTAL] > 0):
        data_found = True
        
        # Iterates through each series count
        for index, count in enumerate(data[date][sheet_name].values()):
          # Writes the date for the current column
          printed_date = date + timedelta(days=1) if (date_shift) else date
          sheet.write(self.date_row, col, printed_date)
          
          # Writes the week number on the 5th row of the current column
          sheet.write(self.week_row, col, col)      # Week number
          
          # Sets write row
          row = self.top_write_row + (2 * index)
          sheet.write(row, col, count)
          
        # Moves to the next column
        col += 1
        
    # Creates charts and adds them to the appropriate locations within sheet
    for chart_param in chart_params:
      self._create_chart(sheet, col - 1, **chart_param)
    
    # Perform final activities on sheet
    self._finalize_sheet(sheet, col - 1)
      
  def _write_data(self, **kwargs):
    """
    Writes the counts and charts to each of the corresponding sheets of the Excel 
    file.
    """
    
    # Creates each sheet with its associated parameters
    for sheet_name, sheet_param in kwargs['sheet_data'].iteritems():
      self._add_sheet(sheet_name, self.data, **sheet_param)