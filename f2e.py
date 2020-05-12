"""

Usage:

  python f2e.py test.csv test.xlsx

"""
import os
import sys
from datetime import datetime,timedelta

do_debug = 'DEBUG' in os.environ

f2e_py_dir = os.path.dirname(__file__)

def hdrcvt(s): return '{1}_{0}'.format(s[:-1].replace('-',''),s[-1])

comma = ','
days1970 = 25569.0
lt_hdrs = 'l1u l1i l2u l2i l3u l3i l4u l4i l4t l1-l2u l2-l3u l3-l1u'.upper().split()
lt_hdr_keys = list(map(hdrcvt,lt_hdrs))
st_hdrs = set(lt_hdrs)


class FREQCSV(object):

  def __init__(self,s_fn_csv):
    self.s_fn_csv = s_fn_csv

  def parse(self):
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

    ### Read row format statement into string
    s_row_xml = os.path.join(f2e_py_dir,'sheet1.xml.row')
    with open(s_row_xml,'r') as frow:
      s_row = frow.read()
      assert not frow.read()

    with open(fn_sheet1,'w') as fsheet:

      ### Copy top of base XML file to sheet XML
      s_top_xml = os.path.join(f2e_py_dir,'sheet1.xml.top')
      with open(s_top_xml,'r') as ftop:
        s_top = ftop.read()
        while s_top:
          fsheet.write(s_top)
          s_top = ftop.read()

      ### Write rows into sheet XML 
      for dt in self.lt_data_rows:
        fsheet.write(s_row.format(**dt))

      ### Copy bottom of base XML file to sheet XML
      s_bottom_xml = os.path.join(f2e_py_dir,'sheet1.xml.bottom')
      with open(s_bottom_xml,'r') as fbottom:
        s_bottom = fbottom.read()
        while s_bottom:
          fsheet.write(s_bottom)
          s_bottom = fbottom.read()


########################################################################
def f2e(s_fn_csv,s_fn_xlsx,s_fn_sheet1='f2e_temp.xml'):

  ### Copy base .ZIP, no_sheet1_xlsx.zip, to XLSX
  s_base_zip = os.path.join(f2e_py_dir,'no_sheet1_xlsx.zip')
  with open(s_base_zip,'rb') as fbase_zip:
    with open(s_fn_xlsx,'wb') as fxlsx:
      data = fbase_zip.read()
      while data:
        fxlsx.write(data)
        data = fbase_zip.read()

  ### Create sheet 1
  fc = FREQCSV(s_fn_csv).parse()
  fc.write_sheet1(s_fn_sheet1)

  ### Add sheet 1 to XLSX ZIP archive file
  import zipfile
  z=zipfile.ZipFile(s_fn_xlsx,'a')
  z.write(s_fn_sheet1,'xl/worksheets/sheet1.xml')
  z.close()

  os.unlink(s_fn_sheet1)



########################################################################
if "__main__" == __name__ and sys.argv[2:]:
  f2e(sys.argv[1],sys.argv[2])
