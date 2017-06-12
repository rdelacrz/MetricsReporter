"""
This module contains classes used to access the back end of the JIRA database
and ClearQuest database.
"""

# Third-party modules
import cx_Oracle
import pyodbc

def log_action(func):
  """
  A decorator function that prints a message when a query is being performed.
  It should be used to wrap any query functions.
  """
  
  def func_wrapper(self, *args, **kwargs):
    """
    Function wrapper that prints the fact that an action is occurring.
    
    @param *args: Arbitrary parameters passed into function.
    @param **kwargs: Arbitrary keyword parameters passed into function.
    """
    
    print 'Querying data...'
    return func(self, *args, **kwargs)
    
  return func_wrapper

class _DBAccessor(object):
  """
  This class represents the general purpose object responsible for logging into
  the back end of various databases. It contains the connect(), disconnect(),
  query(), and query_iterable() functions, and is to be extended by other
  classes in order to account for differences in authentication.

  It should NOT be accessed directly outside of this module.
  """

  def __init__(self, databasemodule, connectstr=''):
    """
    Creates the DBAccessor object. Its initial connection string is empty.

    @param databasemodule: The module for the database being connected to.
    @param connectstr: The connection string used to access the database.
    """

    # Sets the module being used for connect()
    self.database = databasemodule

    # Connection string used in connecting to the database
    self.connectstr = connectstr

    # Connection object that allows access to the database
    self.connection = None

    # Cursor used extract tables within the database
    self.cursor = None

  def connect(self):
    """
    Attempts to establish a connection with the server and returns itself.
    Raises a ConnectionError with an accompanying message if authentication
    fails.
    """

    # Establishes a connection to database if one doesn't already exist
    if (not self.connection):
      self.connection = self.database.connect(self.connectstr)

      # Opens a cursor for queries to use
      self.cursor = self.connection.cursor()

  def disconnect(self):
    """
    Disconnects from the database once no more actions are necessary. This
    function should be called once the user is finished using the database
    to free up its resources.
    """

    # If a connection currently exists, it's closed (along with any cursor)
    try:
      if (self.connection):
        if (self.cursor):
          self.cursor.close()
          self.cursor = None
        self.connection.close()
        self.connection = None
    except Exception, e:
      print "Error message: %s" % str(e)
      self.cursor = None
      self.connection = None

  def query(self, querystr):
    """
    Performs a query with the given string and returns the results as a
    list of tuples. Each item in the list represents a row within the
    resulting table, while the tuple contents represents the cells within
    each row.

    @param querystr: The string of SQL being queried.
    @return: A list of tuples containing the query results. The structure
    of the list is essentially that of a table.
    """

    # Raises exception if no connection to database exists
    if (not self.connection or not self.cursor):
        raise Exception("No connection to database established.")

    # Performs query
    self.cursor.execute(querystr)

    # Obtains list of all results from the query
    table = self.cursor.fetchall()

    # Returns the populated list
    return table

  def query_iterable(self, querystr):
    """
    Performs a query with the given string and returns the cursor, which
    can iterate through each row within the resulting table, while the
    tuple contents represents the cells within each row.

    @param querystr: The string of SQL being queried.
    @return: A cursor object responsible for iterating through each item
    returned from the query.
    """

    # Raises exception if no connection to database exists
    if (not self.connection or not self.cursor):
        raise Exception("No connection to database established.")

    # Performs query
    self.cursor.execute(querystr)

    return self.cursor

####################################################################
#                              Oracle                              #
####################################################################

class Oracle(_DBAccessor):
  """
  This class encapsulates the code responsible for logging into an Oracle
  database.
  """

  def __init__(self, username, password, ip, port, database):
    """
    Initializes the parameters used to authenticate the connection to an
    Oracle database.

    @param username: Username for accessing Oracle database.
    @param password: Password for accessing Oracle database.
    @param ip: IP address of Oracle database.
    @param database: Name of Oracle database.
    """

    self.username = username
    self.password = password
    self.dsn = cx_Oracle.makedsn(ip, port, database)

    super(Oracle, self).__init__(cx_Oracle)

  def connect(self):
    # Establishes a connection to database if one doesn't already exist
    if not self.connection:
      self.connection = self.database.connect(self.username, self.password, self.dsn)

      # Opens a cursor for queries to use
      self.cursor = self.connection.cursor()

####################################################################
#                      Microsoft SQL Server                        #
####################################################################

class MSSQL(_DBAccessor):
  """
  This class encapsulates the code responsible for logging into a MSSQL
  database.
  """

  def __init__(self, username, password, server, database):
    """
    Initializes the parameters used to authenticate the connection to a
    MSSQL database.

    @param username: Username for accessing MSSQL database.
    @param password: Password for accessing MSSQL database.
    @param server: The server on which the database is located.
    @param database: The name of the database on the server.
    """

    # Sets the connection string based on the passed parameters
    connectstr = 'DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s;'\
        % (server, database, username, password)

    # Calls the superclass init() to set appropriate parameters
    super(MSSQL, self).__init__(pyodbc, connectstr)
    
# Imports lower-level class into top level
from jira_gt import JiraGT
from clearquest import ClearQuest