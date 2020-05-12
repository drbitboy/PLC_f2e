#!/usr/bin/env python
#!/usr/bin/env jython?
"""

Usage:

  python f2e.py [test_in.csv [test_out.xlsx]]



"""
import os
import sys
try: import system
except: system = None
from datetime import datetime,timedelta

########################################################################
### Local configuration
########################################################################

def hdrcvt(s):
  """Convert 'l1 u' to U_L1, 'l2-l3 t' to 'T_L2L3', etc."""
  return '{1}_{0}'.format(s[:-1].replace('-',''),s[-1])

### - Destination directory for XLSX product
### - Directory of this file; used to find templates
### - Debugging (not used yet)
xlsx_dest_dir_default = ( os.environ.get('USER','') == 'dad'
                      and '.'
                       or 'C:\\Temp\\janit\\converted'
                        )

f2e_py_dir = os.path.dirname(__file__)

do_debug = 'DEBUG' in os.environ

########################################################################
### - Constants, default headers

comma = ','
days1970 = 25569.0
lt_hdrs = 'l1u l1i l2u l2i l3u l3i l4u l4i l4t l1-l2u l2-l3u l3-l1u'.upper().split()
lt_hdr_keys = list(map(hdrcvt,lt_hdrs))
st_hdrs = set(lt_hdrs)


########################################################################
class FREQCSV(object):
  """Class to parse frequency data in CSV to eXcel .XLSX file"""

  def __init__(self,s_fn_csv):
    """Initialization saves CSV path name"""
    self.s_fn_csv = s_fn_csv

  def parse(self):
    """Parse data from CSV file; return self"""
    time_next,hdr_next = True,True
    self.lt_data_rows = list()
    with open(self.s_fn_csv,'r') as fcsv:
      for rawline in fcsv:
        rawtoks = rawline.strip().split(comma)
        csvs = list(map(lambda s:s.replace(' ','').upper(),rawtoks))

        if time_next:
          ### Parse Event time
          assert 'EVENTTIME'==csvs[0]
          assert '.' in csvs[1]
          time_toks = csvs[1].split('.')
          time_toks[-1] = time_toks[-1][:6]
          s_time = '.'.join(time_toks)
          self.unix_time = (datetime.strptime(s_time,'%Y-%m-%d%H:%M:%S.%f')-datetime(1970,1,1)).total_seconds()
          self.unix_time_ms = int(round(self.unix_time * 1e3))
          self.excel_date = days1970 + (self.unix_time / 86400.0)
          time_next = False
          continue

        if hdr_next:
          ### Parse Headers
          if set(csvs).intersection(st_hdrs) == st_hdrs:
            self.hdrs = csvs
            self.hdr_keys = list(map(hdrcvt,csvs))
            self.L = len(self.hdr_keys)
            hdr_next = False
            sheet_row = 2
            uS = 0
            mod100 = 0
          continue

        dt = dict(zip(self.hdr_keys,rawtoks[:self.L]))
        dt.update(dict(row=sheet_row
                      ,Unix_time=self.unix_time_ms + int(uS / 1000)
                      ,uS=uS
                      ,freq=50
                      ,count=mod100+1
                      ))
        self.lt_data_rows.append(dt)

        sheet_row += 1
        uS += 10000
        mod100 = (mod100 + 1) % 100

    return self

  def write_sheet1(self,fn_sheet1):
    """Write parsed CSV data to XLSX worksheet as XML"""

    ### Read row format statement into string
    s_row_xml = os.path.join(f2e_py_dir,'sheet1_template.xml.row')
    with open(s_row_xml,'r') as frow:
      s_row = frow.read()
      assert not frow.read()

    with open(fn_sheet1,'w') as fsheet:

      ### Copy top of base XML file to worksheet XML
      s_top_xml = os.path.join(f2e_py_dir,'sheet1_template.xml.top')
      with open(s_top_xml,'r') as ftop:
        s_top = ftop.read()
        while s_top:
          fsheet.write(s_top)
          s_top = ftop.read()

      ### Write rows into worksheet XML 
      for dt in self.lt_data_rows:
        fsheet.write(s_row.format(**dt))

      ### Copy bottom of base XML file to worksheet XML
      s_bottom_xml = os.path.join(f2e_py_dir,'sheet1_template.xml.bottom')
      with open(s_bottom_xml,'r') as fbottom:
        s_bottom = fbottom.read()
        while s_bottom:
          fsheet.write(s_bottom)
          s_bottom = fbottom.read()


########################################################################
def f2e(s_fn_csv_arg=None
       ,s_fn_xlsx_arg=None
       ,s_dn_out_arg=xlsx_dest_dir_default
       ,s_fn_sheet1_xml_arg=None
       ):

  ### Copy arguments to local variables
  (s_fn_csv    ,s_dn_out    ,s_fn_xlsx    ,s_fn_sheet1_xml,) = (
   s_fn_csv_arg,s_dn_out_arg,s_fn_xlsx_arg,s_fn_sheet1_xml_arg,)

  ### Prompt for missing input CSV filename
  if not s_fn_csv:
    if None is system:
      sys.stdout.write('Enter source CSV, including any path:  ')
      sys.stdout.flush()
      s_fn_csv = sys.stdin.readline().rstrip()
    else:
      s_fn_csv = system.file.openFile()

    if not (s_fn_csv and os.path.isfile(s_fn_csv)):
      sys.stderr.write('CSV file either not found or not provided:  [{0}]\n'.format(s_fn_csv))
      return

  ### Build XLSX filename
  ### N.B. If input argument is not None or False, it is assumed to
  ###      include the directory 
  if not s_fn_xlsx:
    s_basename = os.path.basename(s_fn_csv)
    toks = s_basename.split('.')
    if 1==len(toks): toks.append('xlsx')
    else           : toks[-1] = 'xlsx'
    s_fn_xlsx = os.path.join(s_dn_out,'.'.join(toks))

  ### Build temporary XML filename - it will be deleted after
  if not s_fn_sheet1_xml:
    s_fn_sheet1_xml = s_fn_xlsx + '.sheet1.xml'

  ### Copy base .ZIP (no_sheet1_xlsx.zip) to XLSX
  s_base_zip = os.path.join(f2e_py_dir,'no_sheet1_xlsx.zip')
  with open(s_base_zip,'rb') as fbase_zip:
    with open(s_fn_xlsx,'wb') as fxlsx:
      data = fbase_zip.read()
      while data:
        fxlsx.write(data)
        data = fbase_zip.read()

  ### Create sheet 1
  FREQCSV(s_fn_csv).parse().write_sheet1(s_fn_sheet1_xml)

  ### Add sheet 1 to XLSX ZIP archive file
  import zipfile
  z=zipfile.ZipFile(s_fn_xlsx,'a')
  z.write(s_fn_sheet1_xml,'xl/worksheets/sheet1.xml')
  z.close()

  ### Delete Sheet1 XML file
  if not do_debug: os.unlink(s_fn_sheet1_xml)


########################################################################
if "__main__" == __name__:
  f2e(s_fn_csv_arg=sys.argv[1:] and sys.argv[1] or None
     ,s_fn_xlsx_arg=sys.argv[2:] and sys.argv[2] or None
     )
