# encoding: utf-8
__author__ = 'zhangyuanxiang'

import time
import sys
import os
import MySQLdb
import cx_Oracle
import MySQLdb.cursors
from xlsxwriter.workbook import Workbook
from MySQLdb.constants import FIELD_TYPE
from optparse import OptionParser
from sys import argv
reload(sys)
sys.setdefaultencoding('utf-8')
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.AL32UTF8'
#os.environ['NLS_LANG']='SIMPLIFIED CHINESE_CHINA.ZHS16GBK'
def get_cli_options():

    parser = OptionParser(usage="usage: python %prog [options]",
                          description="""export data to excel""")

    parser.add_option("-H", "--host",
                      dest="host",
                      default="NULL",
                      metavar="host:port:user:passwd:sid",
                      help="hosportuserpasssid")

    parser.add_option("-F", "--out",
                      dest="outfile",
                      default="null",
                      metavar="outfile",
                      help="output file")
    parser.add_option("-T", "--type",
                      dest="type",
                      default="null",
                      metavar="type",
                      help="O or m ")

    parser.add_option("-S", "--sql",
                      dest="sql",
                      default="select 1;",
                      metavar="sqlstring",
                      help="sqlstring")
    (options, args) = parser.parse_args()

    return options

def get_client(hostport,dbtype):
  if dbtype.lower()=='m':
      try:
         host=hostport.strip().split(':')[0]
         port=hostport.strip().split(':')[1]
         username=hostport.strip().split(':')[2]
         password=hostport.strip().split(':')[3]
         conn = MySQLdb.connect(host=host,
                                port=int(port),
                                user=username,
                                passwd=password,
                                charset='UTF8',
                                cursorclass = MySQLdb.cursors.SSCursor 
                                )
      except  MySQLdb.Error,e:
          print "Error %d:%s"%(e.args[0],e.args[1])
          exit(1)
      return conn
  elif  dbtype.lower()=='o':
     try:
        host=hostport.strip().split(':')[0]
        port=hostport.strip().split(':')[1]
        user=hostport.strip().split(':')[2]
        password=hostport.strip().split(':')[3]
        sid=hostport.strip().split(':')[4]
        dsn_tns =cx_Oracle.makedsn(host,port,sid)
        conn = cx_Oracle.connect(user,password,dsn_tns)
     except cx_Oracle.DatabaseError as e:
              print e
     return conn
  else:
     print "OKOKOKOKOKOK"
def export_data(sql,out):
    count=0
    dt=time.strftime('%Y-%m-%d_%H%M%S',time.localtime(time.time()))
    options = get_cli_options()
    cursor=get_client(options.host,options.type).cursor()
    cursor.execute(sql)
    field_names = [i[0] for i in cursor.description]
    workbook = Workbook(out+"_"+dt+'.xlsx',{'constant_memory': True})
    data_A = workbook.add_worksheet()
    data_B = workbook.add_worksheet()
    data_C = workbook.add_worksheet()
    for c,col in enumerate(field_names):
                 data_A.write(0,c, col)
    rows=cursor.fetchmany(10000)
    while rows:
      for r, row in enumerate(rows):
           r+=1
           count+=1
           if count<1000000: 
              for c, col in enumerate(row): 
                   data_A.write(count,c, col)
           elif count>=1000000 and count<2000000:
               if count==1000000:
                  for c,col in enumerate(field_names):
                      data_B.write(0,c, col)
               else:
                  pass
               r1=(count-999999)
               for c, col in enumerate(row):
                    data_B.write(r1, c, col)
           elif count>=2000000 and count<3000000:
               if count==2000000:
                   for c,col in enumerate(field_names):
                        data_C.write(0,c, col)
               else:
                  pass
               r2=(count-1999999)
               for c, col in enumerate(row):
                  data_C.write(r2, c, col)
           else:
               pass
      rows=cursor.fetchmany(10000)
           
def main() :

    options = get_cli_options()
    export_data(options.sql,options.outfile)

if __name__ == '__main__':
   main()
