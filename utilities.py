"""
This module contains functions that perform general purpose functions.
"""

# Built-in modules
from datetime import datetime, timedelta
import os
import re
import shutil
import time

# User-defined modules
from directories import FILES_DIR

def  get_title(html):
  """
  Obtains the title from the given HTML.
  """
  
  title_re = re.search('<title>\s*(.*)\s*</title>', html)
  title = title_re.group(1).encode('UTF-8', 'ignore')
  return title.rstrip()

def get_top_level_path():
  """
  Gets the absolute path of the directory containing this file - essentially
  the top level directory as long as this utilities.py file remains at the
  top level.

  @return: Absolute path of the top level directory.
  """

  return os.path.dirname(os.path.abspath(__file__))

def create_dirpath(basedir=FILES_DIR, subdirs=[]):
  """
  Creates the directory and its following series of sub directories.
  Returns the absolute path of the resulting directory at the lowest level.

  @param basedir: The lowest level directory which is always created (if it
  doesn't already exist), from which the sub-directories are created as well.
  @param subdirs: The list of subdirectories which are progressively created
  within the base directory. Each following folder within the list is created
  within the previous folder in the list. For example, if subdirs =
  [Raw Data, Projects, SEPTA], it would lead to the creation of subfolders
  within the given base directory with the path "Raw Data\\Projects\\SEPTA".
  @return: The absolute path of the resulting directory at the lowest level.
  """

  # Sets cursor for directory
  currdir = get_top_level_path()
  if (currdir[-4:] == '.exe'):
    currdir = os.path.abspath(os.path.join(currdir, os.pardir))
  currdir = '%s\\%s' % (currdir, basedir)

  # Checks whether base directory exists
  if (not os.path.isdir(currdir)):
    os.mkdir(currdir)

  # Creates the series of sub-directories
  for directory in subdirs:
    currdir += '\\%s' % directory
    if (not os.path.isdir(currdir)):
      os.mkdir(currdir)

  return os.path.abspath(currdir)

def save_file(filename, content, ext, basedir=FILES_DIR, subdirs=[], append=False):
  """
  Creates a file with the given content and extension, and returns the
  absolute path of the created file.

  @param filename: The name of the file being saved.
  @param content: The actual content being saved within the file.
  @param ext: The extension of the file.
  @param basedir: The lowest level directory which is always created (if it
  doesn't already exist), from which the sub-directories are created as well.
  @param subdirs: The list of subdirectories which are progressively created
  within the base directory. Each following folder within the list is created
  within the previous folder in the list. For example, if subdirs =
  [Raw Data, Projects, SEPTA], it would lead to the creation of subfolders
  within the given base directory with the path "Raw Data\\Projects\\SEPTA".
  @param append: Determines whether data is being appended to an existing file
  with the given name, or if any such file is simply overwritten completely.
  """

  # Creates path to the directory
  dirpath = create_dirpath(basedir, subdirs)

  # Fixes extension if . is not included
  if (ext[0] != '.'):
    ext = '.' + ext

  # Creates a file with the given filename containing the given data
  filepath = '%s\\%s' % (dirpath, filename)
  if (not filepath.endswith(ext)):
    filepath += ext
  if (append):
    f = open(filepath, 'a')
  else:
    f = open(filepath, 'wb+')
    f.seek(0)
  f.write(content)
  f.close()

  return filepath

def create_file_log(filename, data):
  """
  Creates a log file within the Logs folder in order to record data for
  testing purposes, and returns the absolute path of the log file.
  """
  
  # Log parameters
  ext = '.txt'                # Log file is a simple text file
  directory = 'Log'           # Name of the folder containing the log file
  
  # Creates the log file within the Log directory
  return save_file(filename, data, ext, subdirs=[directory])
    
def move_old_files(strlist, subdirs):
  """
  Moves all the old files with a string from the given list into the
  given folder.

  NOTE: The 'Files' folder is implicitly determined to be the base directory
  of all the subdirectories, and should not be included within the subdirs
  list.

  @param strlist: A list of strings for which filenames containing any of the
  strings should be moved to the old directory.
  @param subdirs: A list of sub-directories for which old files on the
  project level should be moved to the folder identified by the last string
  on the list. For example, subdirs == ['Metrics','Previous Metrics']
  represents the path '..Metrics/Previous Metrics', and all files will be moved
  to the Previous Metrics sub-folder.
  """

  # Gets the string for the current date
  date = time.strftime("%Y-%m-%d")

  # Establishes directory paths
  directory = create_dirpath(subdirs=subdirs[:-1])
  subdir = create_dirpath(subdirs=subdirs)

  # Traverses through list of files in the directory for folder transferral
  for filename in os.listdir(directory):
    for filestr in strlist:
      if (filestr in filename and 'xlsx' in filename and date not in filename):
        src = '%s/%s' % (directory, filename)
        dst = '%s/%s' % (subdir, filename)
        shutil.move(src, dst)
        break
      
def get_past_year(ascending=True):
  """
  Gets a list of dates representing the last 12 months.
  
  @param ascending: True if order should be ascending, False otherwise.
  @return: List of dates representing past 12 months.
  """
  
  # Gets today's date (with day = 1)
  date = datetime.today().date().replace(day=1)
  
  # Gets the past 12 months, in ascending order
  datelist = []
  for _ in range(12):
    datelist.append(date)
    month = date.month - 1 if (date.month) > 1 else 12
    year = date.year if (date.month) > 1 else date.year - 1
    date = datetime(year, month, 1).date()
    
  # Reverses order to correct it, if ascending order is expected
  if (ascending):
    datelist.reverse()
    
  return datelist
        
def get_extraction_date(date, extractionday):
  """
  Takes the given date and obtains the lowest date greater than or equal to
  it that takes place on the extraction day (integer representing a day of
  the week based on the datetime module).

  @param date: The date for which the next extraction date is being obtained.
  @param extractionday: Day of the week when data is extracted, represented
  by an integer where Mon=0, Tue=1, Wed=2, etc.
  @return: The extraction date associated with the given date. It takes place
  on the day of the week identified by the extraction day integer, and it
  either takes place on the same day as the given date or is after that date
  by no more than seven days.
  """

  # The # of days between the given date and the next extraction date
  day = date.weekday()
  diff = extractionday - day
  if (day > extractionday):
    diff += 7  # Gets next week's date

  # Gets the extraction date associated with the given date
  date = date + timedelta(days=diff)

  # Formats dates to remove time
  date = date.replace(hour=0, minute=0, second=0, microsecond=0)

  return date
        
def _append_historical_date(earliestdate, currdate):
  """
  Recursively appends dates to list so that they are in order, and the dates
  are listed in increments of 7 days. The very first date on the list will
  always be greater than or equal to the earliest date, because the data
  associated with the first date will include any metrics associated with the
  earliest date.

  @param earliestdate: A datetime object representing the earliest
  possible date from the data set.
  @param currdate: A datetime object representing the current date being
  appended to the list of dates during the recursion process.
  @return: A list of dates separated by 7 day intervals leading up to the
  given current date from the earliest date.
  """

  # Day of previous week
  prevdate = currdate - timedelta(days=7)
  if (prevdate.date() < earliestdate.date()):
    return [currdate]
  else:
    # Recursive call to obtain list for previous dates
    datelist = _append_historical_date(earliestdate, prevdate)
    datelist.append(currdate)
    return datelist

def get_historical_dates(earliest, extractionday=4, hastime=True):
  """
  Gets a list of every date to get data for, given the earliest date given.
  Each date is separated by weekly intervals.

  @param earliest: A datetime object representing the earliest possible
  date from the data set.
  @param extractionday: Expected day that data will be extracted. Defaults to
  4, which corresponds to Friday (the end of the working week).
  @param hastime: True if list is to contain datetime objects, false, if the
  list should contain simple date objects with no time information.
  @return: A list of datetime objects representing weekly intervals of dates
  when data is supposed to be measured.
  """

  # Gets the data extraction day of the current week
  date = get_extraction_date(datetime.today(), extractionday)

  # Populates list of dates to collect data, up to the current date
  dates = _append_historical_date(earliest, date)
  
  # Determines whether a list of datetime or date objects are returned
  return dates if (hastime) else [date.date() for date in dates]

def get_str_date(string, regex='.+(\d{4})-(\d{2})-(\d{2}).+'):
  """
  Extracts the date from the given string.
  
  @param string: String with the date to extract.
  @param regex: Regualr expression used to locate the date.
  @return: Extracted date.
  """
  
  # Compiles pattern
  date_re = re.compile(regex)
  
  # Gets date from the string
  date_pattern = date_re.search(string)
  if (date_pattern):
    return datetime(*[int(date_pattern.group(x)) for x in [1, 2, 3]])
  else: None

class AverageAge(object):
  """
  Special class used to calculate an average age (in number of days), based on
  the parameters passed. It stores both the sum (as a timedelta object) and 
  number of items involved, in addition to the calculated average. The purpose
  of this class is to save the age_sum and item_num information in the 
  situation that it will be recalculated using additional items.
  """
  
  def __init__(self, age_sum=timedelta(days=0), item_num=0):
    """
    Initializes age_sum, item_num, and calculates initial average.
    
    @param age_sum: The sum of ages, stored as a timedelta object.
    @param item_num: The number of items involved.
    """
    
    # Sets the initial values
    self.sum = age_sum
    self.num = item_num
    
    # Calculates and sets the average
    self.calculate_average()
    
  def __int__(self):
    """
    Returns this object's average (converted into an int) to serve as its value.
    
    @return: The current object's average.
    """
    
    return int(self.average)
  
  def __str__(self):
    """
    Returns the string representation of this object's average.
    
    @return: The current object's average as a string.
    """
    
    return str(self.average)
  
  def __float__(self):
    """
    Returns the current object's average.
    
    @return: The current object's average.
    """
    
    return float(self.average)
  
  @property
  def value(self):
    """
    Sets the object's average as its value.
    """
    
    return self.average
    
  def calculate_average(self):
    """
    Calculates the average age by dividing the sum of the ages by the number of
    issues, both values stored as instance variables. Sets the result as the
    object's value.
    """
    
    # Divides sum of ages by number issues; rounds to nearest tenth
    if (self.num > 0):
      self.average = round(float(self.sum.total_seconds()) / 86400 / self.num, 1)
    else:
      self.average = 0       # Avoids division by 0
      
  def update(self, value, update_avg=True):
    """
    Adds to the current sum with the given value, and updates the average if
    update_avg is True.
    
    @param value: Value to add to the sum and change the average.
    @param update_avg: True if average is too be updated immediately, False
    otherwise. Setting this value to False increases time efficiency over the
    course of multiple updates, but also means that this object should have its
    average value calculated once all the updates are finished.
    """
    
    self.sum += value
    self.num += 1
    if (update_avg): self.calculate_average()
      
  def combine(self, average_obj, update_avg=True):
    """
    Re-calculates the average age based on the AverageAge object passed as the
    parameter, whose sum and num are to be combined with the current object's.
    
    @param average_obj: Another AverageAge object whose values are being
    combined with the current AverageAge object.
    @param update_avg: True if average is too be updated immediately, False
    otherwise. Setting this value to False increases time efficiency over the
    course of multiple combines, but also means that this object should have its
    average value calculated once all the combines are finished.
    """
    
    self.sum += average_obj.sum
    self.num += average_obj.num
    if (update_avg): self.calculate_average()