'''
This module will be responsible for creating a generic Powerpoint. Subclasses
should extend this module in order to meet more specific needs.
'''

# Built-in modules
from collections import OrderedDict
from datetime import datetime
import os
import re

# Third-party modules
from xlsxwriter.utility import xl_range
import pythoncom
from win32com.client import DispatchEx, constants as win32c

# User-defined modules
from directories import PPT_DIR, TEMPLATE_DIR
from utilities import create_dirpath

# Constant that identifies the default top for inserting an Excel object
TOP_INSERT = 155

class Powerpoint(object):
  """
  Class object encapsulating the Powerpoint generation code.
  """
  
  # Slide dimensions
  slide_width = 720
  slide_height = 540
  
  def __init__(self):
    """
    Initializes the parameters related the Powerpoint file to be generated.
    """
    
    # Applications (initially set to None)
    self.excel = None
    self.ppt = None
    
    # Presentation file
    self.ppt_file = None
    
    # Index to be used when adding slides to the Powerpoint presentation
    self.index = 1
    
    # Index for this Powerpoint's blank page within the template
    self.blank_index = 12
    
    # Integer for template color (orange)
    self.colornum = self.rgb(red=230, green=118, blue=0)
    
    # Paths for template files
    self.template_path = create_dirpath(subdirs=[TEMPLATE_DIR])
    self.iso_logo = os.path.join(self.template_path, 'iso_orange_template.png')
    self.ppt_template = os.path.join(self.template_path, 'ppt_orange_template.potx')
    
    # Path for saved Powerpoint files
    self.save_path = create_dirpath(subdirs=[PPT_DIR])
    
  def setup(self):
    """
    Sets up the initial parameters of the Powerpoint using the Object Model.
    """
    
    # Enables the following code to be threaded
    pythoncom.CoInitialize()                        # @UndefinedVariable
    
    # Attempts to generate the Powerpoint using Microsoft's object model
    try:
      # Opens the Excel and Powerpoint applications invisibly
      self.excel = DispatchEx('Excel.Application')
      self.ppt = DispatchEx('Powerpoint.Application')
      
      # Opens up a new Powerpoint presentation
      self.ppt_file = self.ppt.Presentations.Add(WithWindow=False)
      
      # Applies theme
      self.ppt_file.ApplyTheme(self.ppt_template)
    except Exception, e:
      print "Error during setup: %s" % str(e)
      
  def cleanup(self):
    """
    Saves and closes the Powerpoint file, and closes all applications that were
    opened by COM.
    """
    
    # Closes the Powerpoint presentation
    if (self.ppt_file):
      self.ppt_file.Saved = True
      self.ppt_file.Close()
      self.ppt_file = None
            
    # Quits the Excel and Powerpoint applications
    if (self.excel and self.excel.Workbooks.Count == 0):
      self.excel.Quit()
      self.excel = None
    if (self.ppt and self.ppt.Presentations.Count == 0):
      self.ppt.Quit()
      self.ppt = None
    
    # Marks the end of threading and frees memory resources
    pythoncom.CoUninitialize()                    # @UndefinedVariable
    
  def _add_logo(self, slide):
    """
    Adds logo to the given slide.
    
    @param slide: The slide object to add the logo to.
    """
    
    # Adds ISO logo
    if (not os.path.isfile(self.iso_logo)):
      raise Exception('ISO logo file is missing')           # Means ISO file is missing
    slide.Shapes.AddPicture(self.iso_logo, 0, -1, 20, 485, 165, 50)
    
  def rgb(self, red=255, green=255, blue=255):
    """
    Return RGB value based on the integer values of the red, green, and blue.
    By default, rgb() returns a value equivalent to full white.
    """
    
    # Checks to make sure that all three colors are between 0 and 255
    for color in [red, green, blue]:
      if (color < 0 or color > 255):
        raise ValueError('A color value is outside the allowed range (0-255)')
    
    return (blue << 16) + (green << 8) + red
    
  def add_textframe(self, slide, text, size, colorrgb, bold, **attrs):
    """
    Creates a textframe and set it with the given attributes. Returns the 
    resulting textframe object.
    
    @param slide: The slide object to which a new textframe is being added to.
    @param text: The text of the textframe.
    @param size: The font size of the textframe's text.
    @param colorrgb: The RGB representation of the textframe's text color.
    @param bold: True is text is bold, false otherwise.
    @param **attrs: Keyword argument containing various arguments for the
    textframe object when it is first created.
    @return: The textframe object that was added to the slide.
    """
    
    # Creates initial textframe with attrs inside the slide
    textframe = slide.Shapes.AddTextbox(**attrs).TextFrame
    
    # Adds the text with the appropriate parameters to the textframe
    textframe.TextRange.Text = text
    textframe.TextRange.Font.Bold = bold
    textframe.TextRange.Font.Size = size
    textframe.TextRange.Font.Name = 'Arial'         # Always use Arial at Xerox
    textframe.TextRange.Font.Color.RGB = colorrgb
    
    return textframe
      
  def add_slide(self, layout_index, title=None, custom=True, logo=True, 
                footer='Xerox Proprietary and Confidential Information', 
                log='Slide added.'):
    """
    General method for adding a slide to the Powerpoint presentation. Returns
    the slide that was added.
    
    @param layout_index: The number associated with a given layout.
    @param title: Title to add to the slide, if any.
    @param custom: True if layout is a custom layout, False if it is a regular
    layout built into the Powerpoint application.
    @param logo: True if logo is to be added, false otherwise.
    @param footer: The footer of the slide.
    @param log: Log information to be printed on the screen.
    @return: The slide object that was added to the Powerpoint file.
    """
    
    # Returns custom slide object
    if (custom):
      layout = self.ppt_file.SlideMaster.CustomLayouts(layout_index)
      slide = self.ppt_file.Slides.AddSlide(self.index, layout)
    else:
      slide = self.ppt_file.Slides.Add(self.index, layout_index)
    
    # Adds logo to bottom right of slide
    if (logo): self._add_logo(slide)
    
    # Increments the slide number
    self.index += 1
    
    # Adds a title to the slide, if any
    if (title):
      try:
        slide.Shapes.Title.TextFrame.TextRange.Text = title
      except:
        self.add_textframe(slide, title, 28, self.colornum, True, 
                           Orientation=1, Left=10, Top=10, Width=370, Height=50)
    
    # Insert footer, if it exists
    if (footer):
      slide.Shapes.AddPlaceholder(15)
      slide.HeadersFooters.Footer.Text = footer
    
    # Prints logging information
    if (log): print log
    
    return slide
      
  def add_title_slide(self, title, subtitle=None, title_size=43, layout_index=1):
    """
    Adds title slide to the Powerpoint presentation with the given week ending
    date as its subtitle.
    
    @param title: The title of the Powerpoint presentation.
    @param subtitle: The subtitle of the Powerpoint presentation (optional).
    @param title_size: The font size of the title (defaults to 43).
    @param layout_index: The index for the title slide (defaults to 26).
    @return: Slide object.
    """
    
    # Adds title slide
    title_slide = self.add_slide(layout_index, title=title, custom=False, 
                            logo=True, footer=None, log='Title slide added.')
    
    # Sets title size
    title_slide.Shapes.Title.TextFrame.TextRange.Font.Size = title_size
    
    # Sets subtitle
    if (subtitle):
      title_slide.Shapes.Placeholders(2).TextFrame.TextRange = subtitle
      
    return title_slide
      
  def add_ending_slide(self, layout_index=39):
    """
    Adds ending slide to the Powerpoint presentation, which will contain an
    illustration with the Xerox logo.
    
    @param layout_index: Layout index of the ending slide.
    @return: Slide object.
    """
    
    return self.add_slide(layout_index, custom=True, logo=False, footer=None, 
                   log='Ending slide added.')
      
  def add_secondary_header(self, slide, text, top=85):
    """
    Adds secondary black text below the regular header on the given slide.
    
    @param slide: The Powerpoint slide with content being added.
    @param text: The text to be added below the header.
    @param top: The y-coordinate of the top of the header within the slide.
    """
    
    self.add_textframe(slide, text, 28, 0, True,
                       Orientation=1, Left=24, Top=top, Width=500, Height=50)
      
  def _format_bullets(self, items):
    """
    Formats text for generating the bullets within the items, based on how the
    list of items is structured.
    
    @param items: List of strings (or another list) to be added to the slide, 
    with each individual item being added to a single bullet point. If another
    list is encountered within the list, it is treated as an indentation within
    the existing bullet points.
    @return: The resulting formatted string based on the passed items.
    """
    
    # Base string to append to
    base_str = ''
    
    # Iterates through each item in the list
    for item in items:
      if isinstance(item, str):
        base_str += '%s\n' % item
      else:
        base_str = base_str.strip() + '\r'
        for x in item:
          base_str += '\t%s\n' % x    # Contains an additional indent
        base_str = base_str.strip() + '\r'
        
    return base_str.strip()
      
  def add_bullets_slide(self, title, items, left=10, top=70, width=650, height=50,
            text_size=18, text_color=0, bullet_color=0, log='Bullets slide added'):
    """
    Adds slide with bullet points, based on the list of items passed.
    
    @param title: Title of the slide.
    @param items: List of strings (or another list) to be added to the slide, 
    with each individual item being added to a single bullet point. If another
    list is encountered within the list, it is treated as an indentation within
    the existing bullet points.
    @param left: Distance from the left of the slide where the text frame with
    bullet points will be inserted.
    @param top: Distance from the top of the slide where the text frame with
    bullet points will be inserted.
    @param width: Width of the text frame with bullet points.
    @param height: Height of the text frame with bullet points.
    @param text_size: Font size of the text.
    @param text_color: Color of text.
    @param bullet_color: Color of bullets.
    @param log: Log information to be printed on the screen.
    @return: Slide object.
    """
    
    # Creates initial empty slide
    slide = self.add_slide(self.blank_index, title=title, log=log)
    
    # Gets formatted string for generating the desired bullet points
    bullet_str = self._format_bullets(items)
    
    # Adds text frame to the slide
    textframe = self.add_textframe(slide, bullet_str, text_size, text_color, False, 
                Orientation=1, Left=left, Top=top, Width=width, Height=height)
    
    # Sets bullet point parameters of text frame
    textframe.TextRange.ParagraphFormat.Bullet.Visible = True
    textframe.TextRange.ParagraphFormat.Bullet.Type = 1
    textframe.TextRange.ParagraphFormat.Bullet.RelativeSize = 1
    textframe.TextRange.ParagraphFormat.Bullet.Font.Color = bullet_color
    
    # Sets indent levels
    for para_count in range(1, textframe.Textrange.Paragraphs().Count + 1):
      paragraph = textframe.Textrange.Paragraphs(para_count)
      
      # Paragraphs with indents will have an increased indent level
      tab_count = paragraph.Text.count('\t')
      if (tab_count):
        paragraph.IndentLevel = 4
        for _ in range(tab_count): paragraph.Replace('\t', '')
          
        # Decreases the font of these lines
        paragraph.Font.Size = 14
        paragraph.ParagraphFormat.Bullet.RelativeSize = 0.75
      else:
        paragraph.IndentLevel = 2
        
    return slide
        
  def add_ppt_table(self, slide, table_map, merges=[], **attrs):
    """
    Adds a table to the slide using the given parameters.
    
    @param slide: Slide object to which the table is being added to.
    @param table_map: An ordered data dictionary mapping tuples to lists. Each
    tuple contains a header and column width, and each tuple is mapped to the
    list of items that are supposed to be below its header.
    @param merges: List of tuples for merging cells within the table. Each
    tuple has the following structure:
      (starting row, starting column, ending row, ending column).
    @param **attrs: Attributes to pass into the AddTable function when creating
    the table object. The following parameters are required: NumRows, 
    NumColumns, Left, Top, and Height.
    """
    
    # Creates initial table
    table = slide.Shapes.AddTable(**attrs).Table
    
    # Merges cells
    for merge in merges:
      row1, col1, row2, col2 = merge
      table.Cell(row1, col1).Merge(MergeTo=table.Cell(row2, col2))
    
    # Sets column widths and writes to the table
    for col, ((header, width), items) in enumerate(table_map.iteritems()):
      col += 1    # Normalizes column number
      
      # Set column width
      table.Columns(col).Width = width
      
      # Sets header
      table.Cell(1, col).Shape.TextFrame.TextRange.Text = header
      table.Cell(1, col).Shape.TextFrame.TextRange.Font.Bold = True
      
      # Add all content to current column
      for row, item in enumerate(items):
        table.Cell(2 + row, col).Shape.TextFrame.TextRange.Text = item
        
    return table
      
  def paste_excel_range(self, xl_sheet, ppt_slide, col1, row1, col2, row2,
                        width=None, height=None, left=0, top=TOP_INSERT):
    """
    Pastes the cells with the given range (as a picture) from the given Excel 
    sheet to the given Powerpoint slide at the given coordinates.
    
    @param xl_sheet: The Excel sheet object on which the Copy() is being performed.
    @param ppt_slide: The Powerpoint slide object on which the Paste() is being
    performed.
    @param col1: The first column of Copy within the Excel sheet (0-indexed).
    @param row1: The first row of Copy within the Excel sheet (0-indexed).
    @param col2: The last column of Copy within the Excel sheet (0-indexed).
    @param row2: The last row of Copy within the Excel sheet (0-indexed).
    @param width: The width of the picture within the Powerpoint.
    @param height: The height of the picture within the Powerpoint.
    @param left: The left coordinate from which the picture is pasted.
    @param top: The top coordinate from which the picture is pasted.
    @return: Object reference to the table that was pasted into the slide.
    """
    
    # Copies the given range within Excel sheet
    cells = xl_sheet.Range(xl_range(row1, col1, row2, col2))
    cells.CopyPicture()
    
    # Pastes Copied content from Excel sheet into Powerpoint slide
    shaperange = ppt_slide.Shapes.Paste()
    
    # Modifies pasted item's location and size
    shape = shaperange.Item(1)
    if (width): shape.Width = width
    if (height): shape.Height = height
    shape.Left = left
    shape.Top = top
    
    return shape
    
  def paste_excel_chart(self, xl_sheet, chart_name, ppt_slide, width=680, 
                        height=300, top=TOP_INSERT):
    """
    Obtains the specified chart from the given sheet, performs a Copy on
    it, and pastes it to the given Powerpoint slide, based on the parameters
    passed.
    
    @param xl_sheet: Worksheet of the Excel file with metric charts on it.
    @param chart_name: Name of the chart to be obtained from the sheet.
    @param ppt_slide: The Powerpoint slide object on which the Paste() is being
    performed.
    @param width: Width of chart within the slide.
    @param height: Height of chart within the slide.
    @param top: Distance from top of slide at which chart is inserted.
    @return: The chart object that was placed inside the slide.
    """
    
    # Gets the Chart Object containing the chart
    try:    chartobject = xl_sheet.ChartObjects(chart_name)
    except: chartobject = xl_sheet.ChartObjects(2)      # Gets the second chart
    chartobject.Select()
    
    # Adds border to the chart
    chartobject.Border.LineStyle = 1
    
    # Makes sure axis is shown in increments of seven days (one week)
    chart = chartobject.Chart
    axis = chart.Axes(1)
    try:
      axis.MajorUnit = 7
      axis.MajorUnitScale = 0
      axis.MinorUnit = 1
      axis.MinorUnitScale = 0
    except Exception:
      pass                            # Empty chart has no axis
    
    # Copy chart
    chartobject.Copy()
    
    # Paste the chart into PPT presentation
    shaperange = ppt_slide.Shapes.Paste()
    chart = shaperange.Item(1)
    chart.Width = width
    chart.Height = height
    chart.Left = (self.slide_width - width) / 2
    chart.Top = top
    
    return chart
  
# Imports sub modules and places them on top level
from data_source_ppt import GTJiraPPT, ClearQuestPPT