"""
Contains code for producing various kinds of Excel writers.
"""

# Built-in modules
from time import strftime

# Third-party modules
from utilities import create_dirpath
from xlsxwriter import Workbook

class XLWriter(object):
  """
  Abstract class for producing a generic Excel sheet from scratch. It should be
  implemented by sub-classes that use xlsxwriter to produce Excel sheets.
  """
  
  def __init__(self, filename='Data', savepath=create_dirpath()):
    """
    Initializes basic Excel parameters.
    
    @param filename: Name of the file (without file extension).
    @param savepath: File path where Excel sheet will be saved at.
    """
    
    # Initializes the workbook of the Excel file
    self.filepath = '%s\\%s %s.xlsx' % (savepath, filename, strftime("%Y-%m-%d"))
    self.wb = Workbook(self.filepath)
    
    # Data that will be written to the Excel file
    self.data = None
    
    # Type of workbook data being written
    self.data_type = '<undefined>'
    
  def _set_data(self, data):
    """
    Sets the object's data dictionary to the one passed as a parameter.
    
    @param data: The data being written to the Excel workbook.
    """
    
    self.data = data
    
  def _write_data(self, **kwargs):
    """
    Uses the writer's data to write to the Excel Workbook being created.
    """
    
    raise NotImplemented('Needs to be implemented by subclass.')
  
  def produce_workbook(self, data, **kwargs):
    """
    Using the data that was passed as a parameter, this function writes the
    data to the workbook based on the implementation of write_data(), and saves
    the Excel file.
    
    @param data: The data being written to the Excel workbook.
    @param **kwargs: Contains any other variables necessary for the write.
    @return: The file path of the Excel workbook that was created.
    """
    
    # Sets the data, writes it to the file, and closes it
    print '%s data being exported to a workbook...' % self.data_type
    self._set_data(data)
    self._write_data(**kwargs)
    self.wb.close()
    
    return self.filepath
    
# Makes writers accessible from the top level
from raw_data_writer import RawDataWriter
from export_data_writer import ExportDataWriter
from table_data_writer import TableDataWriter