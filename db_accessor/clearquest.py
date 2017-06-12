"""
This module is responsible for pulling data from ClearQuest, which uses MS SQL 
for its back end.
"""

# Built-in modules
from collections import OrderedDict
from os.path import join

# Third-party modules
import yaml

# User-defined modules
from constants import *
from directories import CONFIG_DIR
from db_accessor import MSSQL, log_action
from utilities import get_top_level_path

class ClearQuest(MSSQL):
  """
  This class encapsulates the code responsible for performing queries on the
  back end of ClearQuest and obtaining its data.
  """
  
  def __init__(self):
    # Gets the ClearQuest parameters
    conn_path = join(get_top_level_path(), CONFIG_DIR, 'connections.yaml')
    conn = yaml.load(open(conn_path, 'r'))[self.__class__.__name__]
    
    # Initializes MSSQL instance using parameters from connections file
    super(ClearQuest, self).__init__(**conn)
    
  def _set_status_history_data(self, data, project, table, issue_type_str, project_str):
    """
    Gets status history data based on the given parameters, and inserts it into
    the historical location of the appropriate project key (within the given data).
    
    It also removes any keys from the main data that has no status transitions,
    since all data in ClearQuest should have status history.
    
    @param data: Data dictionary containing the regular data for each individual key.
    @param project: Name of project.
    @param table: Name of SQL table to pull data from.
    @param issue_type_str: SQL string related to issue types.
    @param project_str: SQL string related to project.
    """
    
    # Query for the change history in the current project
    query = """
    SELECT TMS.cqadmin.%s.id, old_state, new_state, action_timestamp
    FROM TMS.cqadmin.history
      LEFT JOIN TMS.cqadmin.%s ON TMS.cqadmin.history.entity_dbid=TMS.cqadmin.%s.dbid
      %s
    WHERE name='%s' AND %s
    ORDER BY TMS.cqadmin.%s.id, action_timestamp
    """ % (table, table, table, project_str, project, issue_type_str, table)
    
    # Maps everything in change history to data
    for key, old, new, date in self.query_iterable(query):
      if (key in data):
        # Creates historical field
        if (HIST not in data[key]): 
          data[key][HIST] = OrderedDict([(OLD, []), (NEW, []), (TRANS, [])])
          
        # No value is typical of first status
        if (old == 'no_value'): old = None
        else: old = old.replace('_', ' ')
        
        # Removes underscores
        new = new.replace('_', ' ')
        
        # Adds change history
        for field, param in [(OLD, old), (NEW, new), (TRANS, date)]:
          data[key][HIST][field].append(param)
      else:
        del data[key]
          
  @log_action
  def get_scr_data(self, project='', scr_type=[DEFECT, ENHANCE]):
    """
    Gets the Software Change Request (SCR) data for the given project.
    
    @param project: Project key for the project being queried for.
    @param scr_type: SCR type (or list of SCR types) being queried for.
    @return: Data dictionary mapping issue keys to field names to associated
    parameters, containing the SCR data associated with the given project.
    """
    
    # Turns a string into a list
    if (isinstance(scr_type, str)): scr_type = [scr_type]
      
    # Fields for queried data
    fields = [HEADLINE, ISSUETYPE, PRIORITY, STATUS, LINKS, PROPERTY, 
              SUBMIT_DATE, CLOSED_DATE, EST_FIX_TIME, ACT_FIX_TIME]
    
    # Query for the issues in the current project
    query = """
    SELECT TMS.cqadmin.scr.id, headline_2, enhancement_request, severity, 
      TMS.cqadmin.statedef.name, linked, project_property, submitted_on, 
      closed_on, estimated_time_to_fix, actual_time_to_fix 
    FROM TMS.cqadmin.scr
      LEFT JOIN TMS.cqadmin.statedef ON state          = TMS.cqadmin.statedef.id
      JOIN TMS.cqadmin.contract ON found_in_city2 = TMS.cqadmin.contract.dbid
    WHERE contract.name='%s' AND enhancement_request IN ('%s')
    ORDER BY TMS.cqadmin.scr.id
    """ % (project, "','".join(scr_type))
    
    # Populates data using queried information
    data = OrderedDict()
    for results in self.query_iterable(query):
      data[results[0]] = OrderedDict()
      for name, val in zip(fields, results[1:]):
        if (name == STATUS): val = val.replace('_', ' ')
        if (name == PRIORITY and val): val = val.split()[-1]    # Gets priority
        data[results[0]][name] = val
        
    # Pulls historical data and inserts it into data dictionary
    table = 'scr'
    issue_type_str = "enhancement_request IN ('%s')" % "','".join(scr_type)
    project_str = 'JOIN TMS.cqadmin.contract ON found_in_city2=TMS.cqadmin.contract.dbid'
    self._set_status_history_data(data, project, table, issue_type_str, project_str)
    
    return data
  
  @log_action
  def get_rr_data(self, project='', rr_type=[PROD, DEV]):
    """
    Gets the Release Request (RR) data for the given project.
    
    @param project: Project key for the project being queried for.
    @param rr_type: RR type (or list of RR types) being queried for.
    @return: Data dictionary mapping issue keys to field names to associated
    parameters, containing the RR data associated with the given project.
    """
    
    # Turns a string into a list
    if (isinstance(rr_type, str)): rr_type = [rr_type]
      
    # Fields for queried data
    fields = [HEADLINE, REL_TYPE, STATUS, PROPERTY, SUBMIT_DATE, CLOSED_DATE, 
              EST_FIELD_DATE, ACT_FIELD_DATE]
    
    # Query for the issues in the current project
    query = """
    SELECT TMS.cqadmin.releaserequest.id, headline, release_type, 
      TMS.cqadmin.statedef.name, rr_product_property, submit_date, closed_date, 
      est_field_date, actual_field_date 
    FROM TMS.cqadmin.releaserequest
      LEFT JOIN TMS.cqadmin.statedef ON state=TMS.cqadmin.statedef.id
      JOIN TMS.cqadmin.contract ON property=TMS.cqadmin.contract.dbid
    WHERE TMS.cqadmin.contract.name='%s' AND release_type in ('%s')
    ORDER BY TMS.cqadmin.releaserequest.id
    """ % (project, "','".join(rr_type))
    
    # Populates data using queried information
    data = OrderedDict()
    for results in self.query_iterable(query):
      data[results[0]] = OrderedDict()
      for name, val in zip(fields, results[1:]):
        if (name == STATUS): val = val.replace('_', ' ')
        data[results[0]][name] = val
        
    # Pulls historical data and inserts it into data dictionary
    table = 'releaserequest'
    issue_type_str = "release_type IN ('%s')" % "','".join(rr_type)
    project_str = 'JOIN TMS.cqadmin.contract on property=TMS.cqadmin.contract.dbid'
    self._set_status_history_data(data, project, table, issue_type_str, project_str)
    
    return data
  
  @log_action
  def get_dcr_data(self, project='', dcr_type=[ENG_NOTICE, ENG_CHANGE]):
    """
    Gets the Document Change Request (DCR) data for the given project.
    
    @param project: Project key for the project being queried for.
    @param dcr_type: DCR type (or list of RR types) being queried for.
    @return: Data dictionary mapping issue keys to field names to associated
    parameters, containing the DCR data associated with the given project.
    """
    
    # Turns a string into a list
    if (isinstance(dcr_type, str)): dcr_type = [dcr_type]
      
    # Fields for queried data
    fields = [SUBMITTER, REL_TYPE, STATUS, LINKED_DOC, SUBMIT_DATE]
    
    # Query for the issues in the current project
    query = """
    WITH d AS (
      SELECT distinct DCR.id, users.login_name, type_of_request, 
        statedef.name AS state, doc.document_num, DCR.submit_date
      FROM TMS.cqadmin.en_ecn AS DCR 
        INNER JOIN TMS.cqadmin.statedef AS statedef  ON DCR.state = statedef.id
        LEFT JOIN TMS.cqadmin.users     AS users     ON DCR.submitter=users.dbid
        LEFT JOIN TMS.cqadmin.parent_child_links AS links 
          ON DCR.dbid = links.parent_dbid AND 16783608 = links.parent_fielddef_id
        LEFT JOIN TMS.cqadmin.contract as contracttbl ON links.child_dbid = contracttbl.dbid
        LEFT JOIN TMS.cqadmin.parent_child_links AS links2 ON DCR.dbid = links2.parent_dbid 
          AND links2.parent_fielddef_id IN (16783604, 16783592, 16785067)
        LEFT JOIN TMS.cqadmin.document AS doc ON links2.child_dbid=doc.dbid
      WHERE DCR.dbid <> 0 and contracttbl.name='%s'),

    doc_group AS (SELECT id,
        STUFF((SELECT DISTINCT ', ' + document_num from d
        WHERE (id=Results.id)
        FOR XML PATH ('')), 1, 1, '') as doclist
    FROM d Results GROUP BY id)
    
    SELECT DISTINCT d.id, login_name, type_of_request, state, doclist, submit_date from d
      LEFT JOIN doc_group ON d.id=doc_group.id
    WHERE d.type_of_request in ('%s') ORDER BY d.id
    """ % (project, "','".join(dcr_type))
    
    # Populates data using queried information
    data = OrderedDict()
    for results in self.query_iterable(query):
      data[results[0]] = OrderedDict()
      for name, val in zip(fields, results[1:]):
        if (name == STATUS): val = val.replace('_', ' ')
        data[results[0]][name] = val
        
    # Pulls historical data and inserts it into data dictionary
    table = 'en_ecn'
    issue_type_str = "type_of_request IN ('%s')" % "','".join(dcr_type)
    project_str = """
    LEFT JOIN TMS.cqadmin.parent_child_links AS links 
      ON en_ecn.dbid = links.parent_dbid AND 16783608 = links.parent_fielddef_id
    LEFT JOIN TMS.cqadmin.contract AS contracttbl ON links.child_dbid = contracttbl.dbid"""
    self._set_status_history_data(data, project, table, issue_type_str, project_str)
    
    return data