#!/usr/bin/env python
"""

Usage (BASH, typical):

  python f2e.py [test_in.csv [test_out.xlsx]]

Python/Jython:

  import f2e

  f2e('short_test_in.csv')

  f2e('short_test_in.csv','excel_directory/short_test_out.xlsx')

  ### See method f2e docstring/comments for more details


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
  return '%s_%s' % (s[-1],s[:-1].replace('-',''),)

### - Debugging (if envvar DEBUG exists, XML file will not be deleted)
### - Destination directory for XLSX product
### - Directory of this file; used to find templates

do_debug = 'DEBUG' in os.environ

xlsx_dest_dir_default = ( os.environ.get('USER','') == 'dad'
                      and '.'
                       or 'C:\\Temp\\janit\\converted'
                        )

try:
  assert not ('F2E_FORCE_FILE_FAIL' in os.environ)
  f2e_py_dir = os.path.dirname(__file__)
except:
  if do_debug:
    import traceback
    traceback.print_exc()
  f2e_py_dir = None
  import f2e_templates

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

    ### Booleans indicate what is next in the sequence:
    ### - Read Event time (first row);
    ### - Read headers  in a later row, before data
    ### - Read data
    time_next,hdr_next = True,True

    ### Create list of dicts; each dict contains per-row data
    self.lt_data_rows = list()

    ### Open file
    fcsv = open(self.s_fn_csv,'r')

    for rawline in fcsv:
      ### Loop over lines inf file
      rawtoks = rawline.strip().split(comma)
      csvs = list(map(lambda s:s.replace(' ','').upper(),rawtoks))

      if time_next:
        ### Step time_next:  parse event time
        assert 'EVENTTIME'==csvs[0]
        assert '.' in csvs[1]
        time_toks = csvs[1].split('.')
        s_time = '.'.join(time_toks[:-1])
        f_frac = float('.%s' % (time_toks[-1][:6],))
        unix_timedelta = datetime.strptime(s_time,'%Y-%m-%d%H:%M:%S')-datetime(1970,1,1)
        self.unix_time = (unix_timedelta.days * 864e2) + unix_timedelta.seconds + f_frac
        self.unix_time_ms = int(round(self.unix_time * 1e3))
        self.excel_date = days1970 + (self.unix_time / 86400.0)
        ### Event time is complete; clear bool for next sequence step
        time_next = False
        continue

      if hdr_next:
        ### Step hdr_next:  parse Headers; ignore rows until 1 matches
        ###                 the expected column headers in set st_hdrs
        if set(csvs).intersection(st_hdrs) == st_hdrs:
          self.hdrs = csvs
          self.hdr_keys = list(map(hdrcvt,csvs))
          self.L = len(self.hdr_keys)
          ### Header is complete; clear bool for next step,
          hdr_next = False
          ### Initialize per-row incrementing counters
          sheet_row = 2
          uS = 0
          mod100 = 0
        continue

      ### Skip any rows with missing data in any column
      if self.L > len([None for s in rawtoks[:self.L] if s.strip()]):
        continue

      ### Parse data rows into dictionary
      dt = dict(zip(self.hdr_keys,rawtoks[:self.L]))
      dt.update(dict(row=sheet_row
                    ,Unix_time=self.unix_time_ms + int(uS / 1000)
                    ,uS=uS
                    ,freq=dt['F_L1']
                    ,count=mod100+1
                    ))
      self.lt_data_rows.append(dt)

      ### Increment counters
      sheet_row += 1
      uS += 10000
      mod100 = (mod100 + 1) % 100

    ### Close file
    fcsv.close()

    ### Return self so FREQCSV(...).parse().write_sheet1(...) will work
    return self


  def write_sheet1(self,fn_sheet1):
    """Write parsed CSV data to XLSX worksheet as XML"""

    if None is f2e_py_dir:
      s_row = f2e_templates.sheet1_template_xml_row
    else:
      ### Read row format statement into string
      s_row_xml = os.path.join(f2e_py_dir,'sheet1_template.xml.row')
      frow = open(s_row_xml,'r')
      s_row = frow.read()
      assert not frow.read()
      frow.close()

    fsheet = open(fn_sheet1,'w')

    if None is f2e_py_dir:
      fsheet.write(f2e_templates.sheet1_template_xml_top)
    else:
      ### Copy top of base XML file to worksheet XML
      s_top_xml = os.path.join(f2e_py_dir,'sheet1_template.xml.top')
      ftop = open(s_top_xml,'r')
      s_top = ftop.read()
      while s_top:
        fsheet.write(s_top)
        s_top = ftop.read()
      ftop.close()

    ### Write rows into worksheet XML 
    for dt in self.lt_data_rows:
      fsheet.write(s_row % dt)

    if None is f2e_py_dir:
      fsheet.write(f2e_templates.sheet1_template_xml_bottom)
    else:
      ### Copy bottom of base XML file to worksheet XML
      s_bottom_xml = os.path.join(f2e_py_dir,'sheet1_template.xml.bottom')
      fbottom = open(s_bottom_xml,'r')
      s_bottom = fbottom.read()
      while s_bottom:
        fsheet.write(s_bottom)
        s_bottom = fbottom.read()
      fbottom.close()

    fsheet.close()

### End of class FREQCSV
######################################################################

########################################################################
def f2e(s_fn_csv_arg=None
       ,s_fn_xlsx_arg=None
       ,s_dn_out_arg=xlsx_dest_dir_default
       ,s_fn_sheet1_xml_arg=None
       ):
  """
Convert frequency and voltage data from CSV file to eXcel XLSX format

Arguments:

  s_fn_csv_arg        - Path to input CSV file (optional)
                        - If not supplied, the user will be prompted,
                          possibly via GUI file selection dialog (FSD)
  s_fn_xlsx_arg       - Path to output XLSX file (optional)
  s_dn_out_arg        - Directory where XLSX file should be written, if
                        argument s_fn_xlsx_arg was not supplied
                        - Default is [%s]
  s_fn_sheet1_xml_arg - Name of temporary XML file (optional)

  """ % (xlsx_dest_dir_default,)
  ###                      - Default is [%s]
  ###""".format(xlsx_dest_dir_default)

  ######################################################################
  ### Part 1:  input argument handling
  ######################################################################

  ### Copy arguments to local variables
  (s_fn_csv    ,s_dn_out    ,s_fn_xlsx    ,s_fn_sheet1_xml,) = (
   s_fn_csv_arg,s_dn_out_arg,s_fn_xlsx_arg,s_fn_sheet1_xml_arg,)

  ### Prompt for missing input CSV filename
  if not s_fn_csv:

    if None is system:
      ### Command line if system module is not present
      sys.stdout.write('Enter source CSV, including any path:  ')
      sys.stdout.flush()
      s_fn_csv = sys.stdin.readline().rstrip()
    else:
      ### If system module is not present, use FSD
      s_fn_csv = system.file.openFile()

    ### Exit if nothing useful was provided
    if not (s_fn_csv and os.path.isfile(s_fn_csv)):
      sys.stderr.write('CSV file either not found or not provided:  [%s]\n' % (s_fn_csv,))
      return

  ### Build default XLSX filename, if s_fn_xlsx_art  was not provided
  ### N.B. If input argument is provided, it is assumed to include the
  ###      directory 
  if not s_fn_xlsx:
    s_basename = os.path.basename(s_fn_csv)
    toks = s_basename.split('.')
    if 1==len(toks): toks.append('xlsx')
    else           : toks[-1] = 'xlsx'
    s_fn_xlsx = os.path.join(s_dn_out,'.'.join(toks))

  ### Build temporary XML filename - it will be deleted after
  if not s_fn_sheet1_xml:
    s_fn_sheet1_xml = s_fn_xlsx + '.sheet1.xml'

  ######################################################################
  ### Part 2:  parse CSV
  ######################################################################

  ### Create sheet 1
  FREQCSV(s_fn_csv).parse().write_sheet1(s_fn_sheet1_xml)

  ######################################################################
  ### Part 3:  write XLSX
  ######################################################################

  ### Copy base .ZIP (no_sheet1_template_xlsx.zip) to XLSX
  fxlsx = open(s_fn_xlsx,'wb')
  if None is f2e_py_dir:
    fxlsx.write(f2e_templates.no_sheet1_template_xlsx_zip)
  else:
    s_base_zip = os.path.join(f2e_py_dir,'no_sheet1_template_xlsx.zip')
    fbase_zip = open(s_base_zip,'rb')
    data = fbase_zip.read()
    while data:
      fxlsx.write(data)
      data = fbase_zip.read()
    fbase_zip.close()
  fxlsx.close()

  ### Add sheet 1 to XLSX ZIP archive file
  import zipfile
  z=zipfile.ZipFile(s_fn_xlsx,'a')
  z.write(s_fn_sheet1_xml,'xl/worksheets/sheet1.xml')
  z.close()

  ### Delete Sheet1 XML file
  if not do_debug: os.unlink(s_fn_sheet1_xml)

  ### End of method f2e
  ######################################################################


########################################################################
if "__main__" == __name__:
  f2e(s_fn_csv_arg=sys.argv[1:] and sys.argv[1] or None
     ,s_fn_xlsx_arg=sys.argv[2:] and sys.argv[2] or None
     )
