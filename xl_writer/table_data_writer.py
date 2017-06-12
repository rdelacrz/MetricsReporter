"""
This module contains functions that read data and generate Excel workbooks
with data tables based on the passed data.
"""

# Built-in modules
from collections import OrderedDict

# User-built modules
from utilities import create_dirpath
from xl_writer import XLWriter

class TableDataWriter(XLWriter):
  """
  The class responsible for generating the metrics sheets as well as the
  charts associated with them.
  """
  
  def __init__(self, filename, save_path=create_dirpath(), chart_title='<Title>', 
               side_header='<Row Header>', top_header='<Column Header>'):
    """
    Initializes a workbook with the sheets from the given sheet list,
    as well as each of the items from the item list to be counted.
    
    @param filename: Name of the workbook to be saved.
    @param save_path: Location where the file will be saved.
    @param chart_title: The title of the charts within the tables.
    @param side_header: The name of the header for the side list.
    @param top_header: The name of the header for the top list.
    """
    
    # Initializes basic data
    super(TableDataWriter, self).__init__(filename, save_path)
    self.data_type = 'Table Data'
    
    # Sets the common chart title
    self.chart_title = chart_title
    
    # Sets the common header for the side of the tables
    self.sideheader = side_header
    
    # Sets the common header for the top of the tables
    self.topheader = top_header
    
    # Initializes the dictionary of sheets
    self.sheets = OrderedDict()
    
    # Initializes the various cell formats
    self.titleformat = self.wb.add_format(
            {'font_color': 'white', 'font_size': 20, 'font': 'Arial',
             'bold': True, 'border': True, 'bg_color': '#E26B0A', 
             'align': 'center'}
    )
    self.toprightheaderformat = self.wb.add_format(
            {'font_color': 'black', 'font_size': 9, 'font': 'Arial',
             'italic': True, 'align': 'right', 'top': True, 
             'right': True, 'bg_color': '#FCD5B4'}
    )
    self.botleftheaderformat = self.wb.add_format(
            {'font_color': 'black', 'font_size': 9, 'font': 'Arial',
             'italic': True, 'align': 'left', 'bottom': True, 
             'left': True, 'bg_color': '#FCD5B4'}
    )
    self.topleftdiagformat = self.wb.add_format(
            {'font': 'Arial', 'top': True, 'left': True, 
             'diag_type': 2, 'bg_color': '#FCD5B4'}
    )
    self.botrightdiagformat = self.wb.add_format(
            {'font': 'Arial', 'bottom': True, 'right': True, 
             'diag_type': 2, 'bg_color': '#FCD5B4'}
    )
    self.topitemformat = self.wb.add_format(
            {'font_color': 'black', 'font_size': 11, 'font': 'Arial',
             'bold': True, 'border': True, 'bg_color': '#FCD5B4', 
             'align':'center', 'valign': 'top'}
    )
    self.sideitemformat = self.wb.add_format(
            {'font_color': 'black', 'font_size': 11, 'font': 'Arial',
             'bold': True, 'border': True, 'bg_color': '#FCD5B4', 
             'align': 'left'}
    )
    self.regformat = self.wb.add_format(
            {'font_color': 'black', 'font_size': 11, 'font': 'Arial',
             'border': True}
    )
    self.noteformat = self.wb.add_format(
            {'font_color': 'black', 'font_size': 10, 'font': 'Arial'}
    )
          
  def _create_table(self, sheet, sidelist, toplist, startrow=2, startcol=1):
    """
    Creates the initial empty table to have data added to it.
    
    @param sheet: The sheet to contain the table.
    @param sidelist: Contains the list of items for the side of the table.
    @param toplist: Contains the list of items for the top of the table.
    @param startrow: The starting row within the sheet where table cells
    will be created.
    @param startcol: The starting column within the sheet where table cells
    will be created.
    """
    
    # Writes the number of rows and columns of the table on top of the sheet
    note = '# of rows: %d, # of columns: %d' % (len(sidelist) + 3, len(toplist) + 2)
    sheet.write(0, 0, note, self.noteformat)
    
    # Creates table title cell
    title_len = len(toplist) + 1
    sheet.merge_range(startrow, startcol, startrow, startcol + title_len, 
                      self.chart_title, self.titleformat)
    
    # Creates sub-headers
    sheet.write(startrow+1, startcol+1, self.topheader, self.toprightheaderformat)
    sheet.write(startrow+2, startcol, self.sideheader, self.botleftheaderformat)
    
    # Creates diagonal separators
    sheet.write(startrow+1, startcol, '', self.topleftdiagformat)
    sheet.write(startrow+2, startcol+1, '', self.botrightdiagformat)
    
    # Creates top headers
    for index, header in enumerate(toplist):
      sheet.merge_range(startrow+1, startcol+index+2, startrow+2, 
                        startcol+index+2, header, self.topitemformat)
      # Performs resizing if the width of the table is too small for title
      length = len(header)+10 if (len(toplist) < 3) else len(header)+3
      sheet.set_column(startcol+index+2, startcol+index+2, length)
        
    # Creates side headers
    for index, header in enumerate(sidelist):
        sheet.merge_range(startrow+index+3, startcol, startrow+index+3, 
                          startcol+1, header, self.sideitemformat)
    sheet.set_column(startcol, startcol+1, 6)
      
  def _write_to_table(self, sheet, data, start_row=5, start_col=3, write_func=None):
    """
    Writes data to the given sheet, using the start_row and start_col to
    determine where to begin writing the data.
    
    @param sheet: The sheet to contain the table.
    @param data: A multi-level data dictionary with the following
    structure:
       <Top Level> -> <Secondary Level> -> Data Value
    @param start_row: The starting row within the sheet where table cells
    will be created.
    @param start_col: The starting column within the sheet where table cells
    will be created.
    @param write_func: Function to be applied to the written value, if any.
    """
    
    # Traverses through the multi-layered dictionary for data
    for i, top_level in enumerate(data.keys()):
      for j, secondary_level in enumerate(data[top_level].keys()):
        value = data[top_level][secondary_level]
        if (write_func): value = write_func(value)
        sheet.write(start_row+i, start_col+j, value, self.regformat)
      
  def _write_data(self, **kwargs):
    """
    Writes the data for a project or project group, located within the given data 
    dictionary, onto each appropriate sheet on the workbook in the form of tables.
    
    @param data: A multi-level data dictionary with the following
    structure:
       <Sheet Name> -> <Top Level> -> <Secondary Level> -> Data Value
    """
    
    # Creates sheet for each top level parameter and generates table
    for sheet, param_data in self.data.iteritems():
      self.sheets[sheet] = self.wb.add_worksheet(sheet)
      
      # Gets lists
      first_list =  param_data.keys()
      secondary_list = param_data[first_list[0]].keys()
      
      # Creates table for given sheet
      self._create_table(self.sheets[sheet], first_list, secondary_list)
        
      # Writes data to each sheet
      self._write_to_table(self.sheets[sheet], param_data, **kwargs)