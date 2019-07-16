import argparse
from pathlib import Path
import csv
#pypyodbc can be gotten installed from pypi.python.org.  It is easily installed with pip, or conda.
#alternatively, you can download the distribution and manually install with the setup.py provided.
import pypyodbc
# This script was written to update the IBIO_PADS_LIB database INT_BIONUMS after a batch of part
# numbers have been requested.  The input is a CSV file with MANUPARTNUM and INT_BIONUMs.  
# Currently the script is hard coded to use the column numbers that the first use case used,
# but it should be pretty easy to modify the section of the code that reads the CSV file

TABLES = ['CAPACITOR','CONNECTOR','DIODE_LED','HARDWARE','IC','INDUCTOR','MISC','RESISTOR','SWITCH_RELAY','TRANSISTOR','XTAL_OSC']
#TABLES = ['CAPACITOR']

#Connect to DB
conn = pypyodbc.win_connect_mdb('C:\Users\yasirc\Desktop\UpdateLib\TransportPartLibAdditions.mdb')
cursor = conn.cursor()

#service functions
def retrieve_rows_from_mfrpartnum(key,fields):
  """For now the key is assumed to be the manufactor part number, and the value is assumed to be the int-bio number.
  returns: (Status, Description)
  """
  #This SQL statement searches all tables for rows that have the MANUPARTNUM equal to the key.
  sql = "\nUNION ALL\n".join("SELECT '{table}' AS TABLE_NAME, {keyname}, {fieldnames}\n  FROM {table} WHERE {keyname} = ?".format(\
    table=table,\
    keyname=key[0],\
    fieldnames=",".join(fields)) for table in TABLES) + ';'
  #retrieve the results, and put them into an array of dictionaries so that values can be looked up by
  #column name instead of position of column name.  This isn't very optimal, but it should be a very
  #small list.
  return [dict(zip(("TABLE","MANUPARTNUM")+tuple(fields),row)) for row in cursor.execute(sql,(key[1],)*len(TABLES)).fetchall()]


def perform_update(table,key,fields):
  """Performs an update to the Parts database, based on given set of arguments
  @param table: String.  The table to update in the database
  @param key: Tuple. Of the form (Column, Value).  Defines which rows to update
  @param fields: Dict. Fields to update where keys are columns and values are destination values
  """
  #Unpack fields into list to enforce order for parameters
  field_list = fields.items()
  if(len(field_list)==0):
    return
  sql = "UPDATE {table} SET {fields} WHERE {key}=?;".format(\
      table=table,\
      key=key[0],\
      fields= ",".join(["{col}=?".format(col=col) for col,val in field_list])\
      )
  params = tuple([val for col,val in field_list]) + (key[1],)
  print("SQL: {0} Parms: {1}".format(sql,params))
  ret = cursor.execute(sql,params)
  cursor.commit()


class FieldStatus:
  UNKNOWN = "UNKNOWN"
  UPDATED = "UPDATED"
  EMPTY = "EMPTY"
  MISMATCH = "MISMATCH"
  MATCHES = "MATCHES"
  
  def __init__(self,status,comment=""):
    self.status = status
    self.comment = comment
    self.src = None
    self.dst = None


class KeyStatus:
  UNKNOWN = "UNKNOWN"
  DUPLICATES = "DUPLICATES"
  MISSING = "MISSING"
  UNIQUE = "UNIQUE"
  UPDATED = "UPDATED"
  def __init__(self,status,comment=""):
    self.status = status
    self.comment = comment

class PartDbECO:
  UPDATEABLE_FIELDS = (FieldStatus.EMPTY,)
  """Fields that will be updated.  Note that this does not affect validity.  Valid fields that
  aren't updateable will still allow the ECO through, but won't be changed by the ECO"""
  def __init__(self,key,values):
    self.key = key
    self.values = values
    self.valuestatus = dict(zip(values.keys(),(FieldStatus(FieldStatus.UNKNOWN),)*len(values.keys())))
    self.keystatus = (KeyStatus.UNKNOWN)

  def valid_hard(self):
    """Returns true if the hard requirements for an ECO are met which are: The Key is Unique and 
    each field to update is either empty, or the value to write to it matches the value already there"""
    if(self.keystatus.status != KeyStatus.UNIQUE):
      return False
    for field,status in self.valuestatus.items():
      if(status.status in (FieldStatus.UNKNOWN,FieldStatus.MISMATCH)):
        return False
    return True


  def validate_key(self,rows):
    if(len(rows) == 0):
      return KeyStatus(KeyStatus.MISSING);
    elif(len(rows) > 1):
      return KeyStatus(KeyStatus.DUPLICATES,"Tables: {0}".format(",".join([row['TABLE'] for row in rows])))
    else:
      return KeyStatus(KeyStatus.UNIQUE)

  def validate(self):
    assert self.key[0] == "MANUPARTNUM", "Only MANUPARTNUM is supported as a key right now" 
    rows = retrieve_rows_from_mfrpartnum(self.key,self.values.keys())
    self.keystatus = self.validate_key(rows)

    if(self.keystatus.status != KeyStatus.UNIQUE):
      return

    row = rows[0]
    self.table = row['TABLE']

    for column,value in self.values.items():
      existing_value = row[column]
      status = FieldStatus(FieldStatus.UNKNOWN)
      status.dst = existing_value
      status.src = value
      if(existing_value == None):
          status.status = FieldStatus.EMPTY
      else:
        if(existing_value.upper() == value.upper()):
          status.status = FieldStatus.MATCHES
        else:
          status.status = FieldStatus.MISMATCH
      self.valuestatus[column] = status

  def submit(self):
    """Check which fields needs to be updated and performs the update"""
    if(not self.valid_hard()):
      return
    fields_to_update = {k:v for k,v in self.values.items()\
        if (self.valuestatus[k].status in PartDbECO.UPDATEABLE_FIELDS)}
    if(len(fields_to_update)==0):
      return
    perform_update(self.table, self.key, fields_to_update)
    self.keystatus.status = KeyStatus.UPDATED
    for k in fields_to_update.keys():
      s = self.valuestatus[k]
      if(s.status in PartDbECO.UPDATEABLE_FIELDS):
        s.status = FieldStatus.UPDATED



parser = argparse.ArgumentParser(description="A tool to batch update the MSAccess library")
parser.add_argument('--file', dest="in_file", action='store')
parser.add_argument('--commit', action='store_true')
args = parser.parse_args()

#deterime the input file
if(args.in_file == None):
  parser.print_help()
  exit()


in_path = Path(args.in_file)
print("KEY_STATUS,KEY,TABLE,FIELD_STATUS,FIELD,SRC,DST")
with open(str(in_path), "r") as csvfile:
  csvreader_iter = iter(csv.reader(csvfile, dialect='excel'))
  #Remove the head
  next(csvreader_iter)
  for lnum, csv_row in enumerate(csvreader_iter):
    manupartnum = csv_row[3]
    ibionum = csv_row[2]
    eco = PartDbECO(('MANUPARTNUM',manupartnum),{"INT_BIONUM":ibionum, 'DATASHEET':csv_row[5]})
    eco.validate()
    if(args.commit and eco.valid_hard()):
      eco.submit()
    print("{status},{key},{table},{comment}".format(\
      status=eco.keystatus.status,\
      key=eco.key[1],\
      table=getattr(eco,'table',None),\
      comment=eco.keystatus.comment))
    for field, value in eco.values.items():
      status = eco.valuestatus[field]
      print(',,,{status},{field},"{src}","{dst}"'.format(\
        status=status.status,\
        field=field,\
        src=status.src,\
        dst=status.dst))
conn.close()


