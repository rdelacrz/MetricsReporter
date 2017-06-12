"""
This module is used for creating the Excel file that will contain raw data. 
Contains a RawDataWriter class which encapsulates all the data and functions 
responsible for writing the Excel tables for this information.
"""

# Built-in modules
from collections import OrderedDict
import datetime

# User-defined modules
from utilities import create_dirpath
from xl_writer import XLWriter

class RawDataWriter(XLWriter):
  """
  This class is responsible for reading raw data and creating a table that
  contains all that information.
  """
  
  def __init__(self, sheet_names, header_lists, filename='Raw Data', save_path=create_dirpath()):
    """
    Initializes the Excel file containing raw data, which will be saved at
    the given path with the given file name.
    
    @param sheet_names: A list of sheet names for each sheet being created.
    @param header_lists: A list of lists of tuples, with each header list
    corresponding with a different sheet. It has  the following format: 
      [[(header, column_width)], [(header, column_width)]]
    @param filename: Name of the file (without file extension).
    @param save_path: File path where Excel sheet will be saved at.
    """
    
    super(RawDataWriter, self).__init__(filename, save_path)
    self.data_type = 'Raw'
    
    # Initializes the worksheets of the workbook
    self.sheets = OrderedDict()
    self.row = OrderedDict()
    for name in sheet_names:
      self.sheets[name] = self.wb.add_worksheet(name)
      self.row[name] = 1        # Sets starting row of writes (0-indexed)
    
    # Initializes tables for each sheet
    for index, sheet_name in enumerate(sheet_names):
      headers = header_lists[index]
      self._create_initial_tables(self.sheets[sheet_name], headers)
      
  def _create_initial_tables(self, sheet, headers):
      """
      Creates the initial empty table for the Excel file.
      
      @param sheet: The sheet being initialized.
      @param headers: The list of headers being added to the sheet.
      """
      
      # Creates the header format
      hformat = self.wb.add_format(
        {'bold' : True,
         'font_color' : 'white',
         'font_size' : 10,
         'font' : 'Arial',
         'bg_color' : 'black',
        }
      )
      
      # Creates the headers within both sheets of the Excel file
      for index, (header, width) in enumerate(headers):
        sheet.write(0, index, header, hformat)
        sheet.set_column(index, index, width)
      
  def _add_row(self, data_list, sheet_name=None):
    """
    Adds the list of data items to the sheet at the current write row. The data
    items will be written to the Excel sheet in the order
    
    @param data_list: List of data items.
    @param sheet_name: The name of the sheet being appended to.
    """
    
    # Gets sheet with the given name, or first sheet if name isn't specified
    if (sheet_name):
      sheet = self.sheets[sheet_name]
    else:
      sheet = self.sheets[self.sheets.keys()[0]]
    
    # Adds the issue data to the current row
    for col, data in enumerate(data_list):
      # Cell format (may change for dates)
      cformat = self.wb.add_format({
        'text_wrap' : True, 'font_size' : 10, 'font' : 'Arial'
      })
      
      # Formats dates
      if (data != None and isinstance(data, datetime.datetime)):
        cformat.set_num_format('yyyy-mm-dd')
      elif (data != None and isinstance(data, str)):
        try:
          data = ' '.join(data.encode('ascii', 'ignore').split()).strip()
        except:
          data = 'Error: Was unable to process string'
      elif (not data):
        data = ""
      else:
        pass
      
      # Writes data at the given cell
      sheet.write(self.row[sheet.name], col, data, cformat)
      
    # Increments the current row by one
    self.row[sheet.name] += 1
      
  def _write_data(self, **kwargs):
    """
    Writes the data to the rows of the Excel workbook, one by one.
    """
    
    # Iterates through each of the items within the data
    for sheet_name, data in self.data.iteritems():
      for items in data:
        self._add_row(items, sheet_name)