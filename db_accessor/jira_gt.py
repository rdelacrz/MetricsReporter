"""
This module is responsible for pulling data from the Germantown instance of
JIRA, which uses Oracle for its back end.
"""

# Built-in modules
from collections import OrderedDict
from os.path import join

# Third-party modules
import yaml

# User-defined modules
from constants import *
from directories import CONFIG_DIR
from db_accessor import Oracle, log_action
from utilities import get_top_level_path

class JiraGT(Oracle):
  """
  This class encapsulates the code responsible for performing queries on the
  back end of the JIRA instance in Germantown and obtaining its data.
  """
  
  def __init__(self):
    # Gets the JIRA parameters
    try:
      conn_path = join(get_top_level_path(), CONFIG_DIR, 'connections.yaml')
      conn = yaml.load(open(conn_path, 'r'))[self.__class__.__name__]
    except:
      # Uses hardcoded info in case of failure
      conn = { 'username': 'JIRA513COPY', 'password': 'welcome123', 
              'ip': '10.36.194.121','port': 1521, 'database': 'jiradev'}
    
    # Initializes Oracle instance using parameters from connections file
    super(JiraGT, self).__init__(**conn)
    
  @log_action
  def get_issue_data(self, project='', issuetype=[DEFECT, DEFECT_SUB]):
    """
    Gets the issue data for the given project with the given issue type(s).
    
    @param project: Project key for the project being queried for.
    @param issuetype: Issue type (or list of issue types) being queried for.
    @return: Data dictionary mapping issue keys to field names to associated
    parameters, containing the data associated with the given project.
    """
    
    # Turns a string into a list
    if (isinstance(issuetype, str)):
      issuetype = [issuetype]
      
    # Fields for queried data
    fields = [ISSUETYPE, PRIORITY, STATUS, CREATED, RESOLVED, COMPS, LINKS, 
              PACK, FOUND, PBI, ROOT, DEV_EST, DEV_ACT]
      
    params = (project, "','".join(issuetype))
    query = """
    WITH project_set AS (
      SELECT jiraissue.id FROM jiraissue
      JOIN project   ON project.id=jiraissue.project AND project.pkey IN ('%s')
      JOIN issuetype ON issuetype.id=jiraissue.issuetype AND issuetype.pname IN ('%s')
    ),
  
    est AS ( 
        SELECT jiraissue.id, val.numbervalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='Dev Estimate'
      ),
      
      act AS (
        SELECT jiraissue.id, val.numbervalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='Dev Actuals'
      ),
      
      foundin AS (
        SELECT jiraissue.id, coption.customvalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
          LEFT JOIN customfieldoption coption ON coption.id=val.stringvalue
        WHERE field.cfname='FoundIn'
      ),
      
      rc AS (
        SELECT jiraissue.id, val.textvalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='RootCause'
      ),
      
      pbi AS (
        SELECT jiraissue.id, val.textvalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='PBI_Code'
      ),
      
      complink AS (SELECT SINK_NODE_ID, SOURCE_NODE_ID FROM nodeassociation
      WHERE ASSOCIATION_TYPE='IssueComponent'), 
      
      ungroupedcomp AS (
        SELECT jiraissue.id, component.cname FROM jiraissue
          LEFT JOIN complink      ON jiraissue.id=complink.SOURCE_NODE_ID
          LEFT JOIN component     ON complink.SINK_NODE_ID=component.id
          JOIN project_set        ON project_set.id=jiraissue.id
      ),
      
      withcomps AS (
        SELECT id,
               LTRIM(MAX(SYS_CONNECT_BY_PATH(cname,'| '))
               KEEP (DENSE_RANK LAST ORDER BY curr),'| ') AS comps
        FROM   (SELECT id, cname,
                   ROW_NUMBER() OVER (PARTITION BY id ORDER BY cname) AS curr,
                   ROW_NUMBER() OVER (PARTITION BY id ORDER BY cname) -1 AS prev
                FROM ungroupedcomp
        )
        GROUP BY id
        CONNECT BY prev = PRIOR curr AND id = PRIOR id
        START WITH curr = 1
      ),
      
      jiraissue2 AS (
        SELECT jiraissue.id, project.pkey || '-' || jiraissue.issuenum AS key
        FROM jiraissue
        JOIN project ON project=project.id
      ),
      
      ungroupedlink AS (
        SELECT jiraissue.id, jiraissue2.key FROM jiraissue
          JOIN issuelink    ON jiraissue.id=issuelink.destination
          JOIN jiraissue2   ON issuelink.source=jiraissue2.id
          JOIN project_set  ON project_set.id=jiraissue.id
        UNION
        SELECT jiraissue.id, jiraissue2.key FROM jiraissue
          JOIN issuelink    ON jiraissue.id=issuelink.source
          JOIN jiraissue2   ON issuelink.source=jiraissue2.id
          JOIN project_set  ON project_set.id=jiraissue.id
      ),
      
      withlinks AS (
        SELECT id,
               LTRIM(MAX(SYS_CONNECT_BY_PATH(key,', '))
               KEEP (DENSE_RANK LAST ORDER BY curr),', ') AS links
        FROM   (SELECT id, key,
                   ROW_NUMBER() OVER (PARTITION BY id ORDER BY key) AS curr,
                   ROW_NUMBER() OVER (PARTITION BY id ORDER BY key) -1 AS prev
                FROM ungroupedlink
        )
        GROUP BY id
        CONNECT BY prev = PRIOR curr AND id = PRIOR id
        START WITH curr = 1
      )
      
      SELECT project.pkey || '-' || jiraissue.issuenum, issuetype.pname, priority.pname,
        issuestatus.pname, created, resolutiondate, comps, links, foption.customvalue, 
        foundin.val, pbi.val, rc.val, est.val, act.val
      FROM jiraissue
        LEFT JOIN customfieldvalue  value  ON value.issue=jiraissue.id 
                                           AND value.customfield=14300
        LEFT JOIN customfieldoption foption ON foption.id=value.stringvalue
        LEFT JOIN issuestatus              ON issuestatus.id=jiraissue.issuestatus
        LEFT JOIN priority                 ON priority.id=jiraissue.priority
        LEFT JOIN withcomps                ON withcomps.id=jiraissue.id
        LEFT JOIN withlinks                ON withlinks.id=jiraissue.id
        LEFT JOIN foundin                  ON foundin.id=jiraissue.id
        LEFT JOIN pbi                      ON pbi.id=jiraissue.id
        LEFT JOIN rc                       ON rc.id=jiraissue.id
        LEFT JOIN est                      ON est.id=jiraissue.id
        LEFT JOIN act                      ON act.id=jiraissue.id
        JOIN project_set                   ON project_set.id=jiraissue.id
        JOIN project                       ON project.id=jiraissue.project
        JOIN issuetype                     ON issuetype.id=jiraissue.issuetype
    ORDER BY issuenum
    """ % params
      
    # Populates data dictionary with queried information
    data = OrderedDict()
    for results in self.query_iterable(query):
      data[results[0]] = OrderedDict()
      for name, val in zip(fields, results[1:]):
        if (name in [COMPS, LINKS] and val == None): val = ''
        if (name == COMPS): val = val.replace('|', ',')
        if (name in [PBI, ROOT] and val is not None): val = val.read()
        data[results[0]][name] = val
        
    # Gets the historical data separately (saves querying time)
    query = """
    SELECT  project.pkey || '-' || jiraissue.issuenum, to_char(oldstring) old,
      to_char(newstring) new, changegroup.created changedate 
    FROM jiraissue
      JOIN changegroup    ON issueid=jiraissue.id
      JOIN changeitem     ON changeitem.groupid=changegroup.id AND field='status'
      JOIN project        ON project.id=project AND project.pkey='%s'
      JOIN issuetype      ON issuetype.id=issuetype AND issuetype.pname IN ('%s')
    ORDER BY issuenum, changedate
    """ % (project, "','".join(issuetype))
    
    # Iterates through issue history
    for key, old, new, changedate in self.query_iterable(query):
      if (HIST not in data[key]):
        data[key][HIST] = OrderedDict([(OLD, []), (NEW, []), (TRANS, [])])
      # Adds each value to the proper list
      for field, value in [(OLD, old), (NEW, new), (TRANS, changedate)]:
        data[key][HIST][field].append(value)
        
    return data
  
  @log_action
  def get_roche_issue_data(self, project='', issuetype=[DEFECT, DEFECT_SUB]):
    """
    Gets the issue data for the given project with the given issue type(s) (components broken down per ticket).
    
    @param project: Project key for the project being queried for.
    @param issuetype: Issue type (or list of issue types) being queried for.
    @return: Data dictionary mapping issue keys to field names to associated
    parameters, containing the data associated with the given project.
    """
    
    # Turns a string into a list
    if (isinstance(issuetype, str)):
      issuetype = [issuetype]
      
    # Fields for queried data
    fields = ["Row Number", ISSUETYPE, PRIORITY, STATUS, CREATED, RESOLVED, COMPS, LINKS, 
              PACK, FOUND, PBI, ROOT, DEV_EST, DEV_ACT]
      
    params = (project, "','".join(issuetype))
    query = """
    WITH project_set AS (
      SELECT jiraissue.id FROM jiraissue
      JOIN project   ON project.id=jiraissue.project AND project.pkey IN ('%s')
      JOIN issuetype ON issuetype.id=jiraissue.issuetype AND issuetype.pname IN ('%s')
    ),
  
    est AS ( 
        SELECT jiraissue.id, val.numbervalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='Dev Estimate'
      ),
      
      act AS (
        SELECT jiraissue.id, val.numbervalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='Dev Actuals'
      ),
      
      foundin AS (
        SELECT jiraissue.id, coption.customvalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
          LEFT JOIN customfieldoption coption ON coption.id=val.stringvalue
        WHERE field.cfname='FoundIn'
      ),
      
      rc AS (
        SELECT jiraissue.id, val.textvalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='RootCause'
      ),
      
      pbi AS (
        SELECT jiraissue.id, val.textvalue AS val FROM jiraissue
          JOIN project_set               ON project_set.id=jiraissue.id
          LEFT JOIN customfieldvalue val ON val.issue=jiraissue.id
          LEFT JOIN customfield field    ON field.id=val.customfield
        WHERE field.cfname='PBI_Code'
      ),
      
      complink AS (SELECT SINK_NODE_ID, SOURCE_NODE_ID FROM nodeassociation
      WHERE ASSOCIATION_TYPE='IssueComponent'), 
      
      withcomps AS (
        SELECT jiraissue.id, component.cname comps FROM jiraissue
          LEFT JOIN complink      ON jiraissue.id=complink.SOURCE_NODE_ID
          LEFT JOIN component     ON complink.SINK_NODE_ID=component.id
          JOIN project_set        ON project_set.id=jiraissue.id
      ),
      
      jiraissue2 AS (
        SELECT jiraissue.id, project.pkey || '-' || jiraissue.issuenum AS key
        FROM jiraissue
        JOIN project ON project=project.id
      ),
      
      ungroupedlink AS (
        SELECT jiraissue.id, jiraissue2.key FROM jiraissue
          JOIN issuelink    ON jiraissue.id=issuelink.destination
          JOIN jiraissue2   ON issuelink.source=jiraissue2.id
          JOIN project_set  ON project_set.id=jiraissue.id
        UNION
        SELECT jiraissue.id, jiraissue2.key FROM jiraissue
          JOIN issuelink    ON jiraissue.id=issuelink.source
          JOIN jiraissue2   ON issuelink.source=jiraissue2.id
          JOIN project_set  ON project_set.id=jiraissue.id
      ),
      
      withlinks AS (
        SELECT id,
               LTRIM(MAX(SYS_CONNECT_BY_PATH(key,', '))
               KEEP (DENSE_RANK LAST ORDER BY curr),', ') AS links
        FROM   (SELECT id, key,
                   ROW_NUMBER() OVER (PARTITION BY id ORDER BY key) AS curr,
                   ROW_NUMBER() OVER (PARTITION BY id ORDER BY key) -1 AS prev
                FROM ungroupedlink
        )
        GROUP BY id
        CONNECT BY prev = PRIOR curr AND id = PRIOR id
        START WITH curr = 1
      )
      
      SELECT project.pkey || '-' || jiraissue.issuenum, issuetype.pname, priority.pname,
        issuestatus.pname, created, resolutiondate, comps, links, foption.customvalue, 
        foundin.val, pbi.val, rc.val, est.val, act.val
      FROM jiraissue
        LEFT JOIN customfieldvalue  value  ON value.issue=jiraissue.id 
                                           AND value.customfield=14300
        LEFT JOIN customfieldoption foption ON foption.id=value.stringvalue
        LEFT JOIN issuestatus              ON issuestatus.id=jiraissue.issuestatus
        LEFT JOIN priority                 ON priority.id=jiraissue.priority
        LEFT JOIN withcomps                ON withcomps.id=jiraissue.id
        LEFT JOIN withlinks                ON withlinks.id=jiraissue.id
        LEFT JOIN foundin                  ON foundin.id=jiraissue.id
        LEFT JOIN pbi                      ON pbi.id=jiraissue.id
        LEFT JOIN rc                       ON rc.id=jiraissue.id
        LEFT JOIN est                      ON est.id=jiraissue.id
        LEFT JOIN act                      ON act.id=jiraissue.id
        JOIN project_set                   ON project_set.id=jiraissue.id
        JOIN project                       ON project.id=jiraissue.project
        JOIN issuetype                     ON issuetype.id=jiraissue.issuetype
    ORDER BY issuenum
    """ % params
      
    # Populates data dictionary with queried information
    data = OrderedDict()
    for results in self.query_iterable(query):
      data[results[0]] = OrderedDict()
      for name, val in zip(fields, results[1:]):
        if (name in [COMPS, LINKS] and val == None): val = ''
        if (name == COMPS): val = val.replace('|', ',')
        if (name in [PBI, ROOT] and val is not None): val = val.read()
        data[results[0]][name] = val
        
    # Gets the historical data separately (saves querying time)
    query = """
    SELECT  project.pkey || '-' || jiraissue.issuenum, to_char(oldstring) old,
      to_char(newstring) new, changegroup.created changedate 
    FROM jiraissue
      JOIN changegroup    ON issueid=jiraissue.id
      JOIN changeitem     ON changeitem.groupid=changegroup.id AND field='status'
      JOIN project        ON project.id=project AND project.pkey='%s'
      JOIN issuetype      ON issuetype.id=issuetype AND issuetype.pname IN ('%s')
    ORDER BY issuenum, changedate
    """ % (project, "','".join(issuetype))
    
    # Iterates through issue history
    for key, old, new, changedate in self.query_iterable(query):
      if (HIST not in data[key]):
        data[key][HIST] = OrderedDict([(OLD, []), (NEW, []), (TRANS, [])])
      # Adds each value to the proper list
      for field, value in [(OLD, old), (NEW, new), (TRANS, changedate)]:
        data[key][HIST][field].append(value)
        
    return data
  
  @log_action
  def get_compliance_data(self):
    """
    Gets compliance data.
    
    @return: Data dictionary mapping issue keys to field names to associated
    parameters, containing issues with the Compliance issue type.
    """
      
    # Fields for queried data
    fields = [PROJECT, ISSUETYPE, PRIORITY, STATUS, CREATED, RESOLVED, COMPS, LINKS, 
              PACK, DEV_EST, DEV_ACT]
      
    query = """
    WITH complink AS (SELECT SINK_NODE_ID, SOURCE_NODE_ID FROM nodeassociation
    WHERE ASSOCIATION_TYPE='IssueComponent'), 
        
    ungrouped AS (
      SELECT jiraissue.id, component.cname FROM jiraissue
      LEFT JOIN complink     ON jiraissue.id=complink.SOURCE_NODE_ID
      LEFT JOIN component   ON complink.SINK_NODE_ID=component.id
      JOIN issuetype          ON issuetype.id=jiraissue.issuetype
      WHERE issuetype.pname='Compliance'
    ),
    
    withcomps AS (
      SELECT id,
             LTRIM(MAX(SYS_CONNECT_BY_PATH(cname,', '))
             KEEP (DENSE_RANK LAST ORDER BY curr),', ') AS comps
      FROM   (SELECT id, cname,
                 ROW_NUMBER() OVER (PARTITION BY id ORDER BY cname) AS curr,
                 ROW_NUMBER() OVER (PARTITION BY id ORDER BY cname) -1 AS prev
              FROM ungrouped
      )
      GROUP BY id
      CONNECT BY prev = PRIOR curr AND id = PRIOR id
      START WITH curr = 1
    ),
        
    jiraissue2 AS (
        SELECT jiraissue.id, project.pkey || '-' || jiraissue.issuenum AS key
        FROM jiraissue
        JOIN project on project=project.id
    ),
    
    withlinks AS (
      SELECT jiraissue.id, COLLECT(jiraissue2.key) AS links FROM jiraissue
      JOIN issuelink    ON jiraissue.id=issuelink.destination
      JOIN jiraissue2   ON issuelink.source=jiraissue2.id
      GROUP BY jiraissue.id
    )
    
    SELECT project.pkey || '-' || jiraissue.issuenum, project.pkey, issuetype.pname, priority.pname,
      issuestatus.pname, created, resolutiondate, comps, links, foption.customvalue
    FROM jiraissue
      LEFT JOIN customfieldvalue  value   ON value.issue=jiraissue.id 
                                          AND value.customfield=14300
      LEFT JOIN customfieldoption foption ON foption.id=value.stringvalue
      LEFT JOIN issuestatus               ON issuestatus.id=jiraissue.issuestatus
      LEFT JOIN priority                  ON priority.id=jiraissue.priority
      LEFT JOIN withcomps                 ON withcomps.id=jiraissue.id
      LEFT JOIN withlinks                 ON withlinks.id=jiraissue.id
      JOIN project                        ON project.id=jiraissue.project
      JOIN issuetype                      ON issuetype.id=jiraissue.issuetype
    WHERE issuetype.pname='Compliance'
    ORDER BY issuenum
    """
      
    data = OrderedDict()
    for results in self.query_iterable(query):
      data[results[0]] = OrderedDict()
      for name, val in zip(fields, results[1:]):
        if (name in [LINKS]):
          val = ', '.join(val) if (val) else ''   # List > str
        data[results[0]][name] = val
        
    # Gets the historical data separately (saves querying time)
    query = """
    SELECT  project.pkey || '-' || jiraissue.issuenum, to_char(oldstring) old,
      to_char(newstring) new, changegroup.created changedate 
    FROM jiraissue
      JOIN changegroup    ON issueid=jiraissue.id
      JOIN changeitem     ON changeitem.groupid=changegroup.id AND field='status'
      JOIN project        ON project.id=project
      JOIN issuetype      ON issuetype.id=issuetype AND issuetype.pname='Compliance'
    ORDER BY issuenum, changedate
    """
    
    # Iterates through issue history
    for key, old, new, changedate in self.query_iterable(query):
      if (HIST not in data[key]):
        data[key][HIST] = OrderedDict([(OLD, []), (NEW, []), (TRANS, [])])
      # Adds each value to the proper list
      for field, value in [(OLD, old), (NEW, new), (TRANS, changedate)]:
        data[key][HIST][field].append(value)
        
    return data
  
  def get_active_projects(self):
    """
    Gets list of currently active projects and corresponding project groups.
    Will ignore projects that are archived or haven't been updated since the
    beginning of the previous year.
    
    @return: Data dictionary mapping project group to a list of tuples, where
    each tuple contains the JIRA key and full name of a given project.
    """
    
    query = """
    SELECT DISTINCT pname, pkey, cname FROM project
      JOIN nodeassociation ON SOURCE_NODE_ID=project.id
      JOIN projectcategory ON SINK_NODE_ID=projectcategory.id
    WHERE ASSOCIATION_TYPE='ProjectCategory' AND cname != 'ARCHIVE'
    ORDER BY cname, pkey"""
    
    data = OrderedDict()
    for group, key, name in self.query_iterable(query):
      if group not in data: data[group] = []
      data[group].append((key, name))
    return data