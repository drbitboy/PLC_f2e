sheet1_template_xml_bottom="""  </sheetData>
  <printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.511805555555555" footer="0.511805555555555"/>
  <pageSetup paperSize="9" scale="100" firstPageNumber="0" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="false" horizontalDpi="300" verticalDpi="300" copies="1"/>
  <headerFooter differentFirst="false" differentOddEven="false">
    <oddHeader/>
    <oddFooter/>
  </headerFooter>
</worksheet>
"""
sheet1_template_xml_row="""    <row r="%(row)s" customFormat="false" ht="15" hidden="false" customHeight="false" outlineLevel="0" collapsed="false">
      <c r="A%(row)s" s="8" t="n">
        <f aca="false">(B%(row)s/86400/1000)+DATE(1970,1,1)</f>
      </c>
      <c r="B%(row)s" s="1" t="n">
        <v>%(Unix_time)s</v>
      </c>
      <c r="C%(row)s" s="1" t="n">
        <f aca="false">D%(row)s/1000</f>
        <v>0</v>
      </c>
      <c r="D%(row)s" s="9" t="n">
        <v>%(uS)s</v>
      </c>
      <c r="E%(row)s" s="9" t="n">
        <v>%(U_L1)s</v>
      </c>
      <c r="F%(row)s" s="9" t="n">
        <v>%(I_L1)s</v>
      </c>
      <c r="G%(row)s" s="9" t="n">
        <v>%(U_L2)s</v>
      </c>
      <c r="H%(row)s" s="9" t="n">
        <v>%(I_L2)s</v>
      </c>
      <c r="I%(row)s" s="9" t="n">
        <v>%(U_L3)s</v>
      </c>
      <c r="J%(row)s" s="9" t="n">
        <v>%(I_L3)s</v>
      </c>
      <c r="K%(row)s" s="9" t="n">
        <v>%(U_L4)s</v>
      </c>
      <c r="L%(row)s" s="9" t="n">
        <v>%(I_L4)s</v>
      </c>
      <c r="M%(row)s" s="9" t="n">
        <v>%(U_L1L2)s</v>
      </c>
      <c r="N%(row)s" s="9" t="n">
        <v>%(U_L2L3)s</v>
      </c>
      <c r="O%(row)s" s="9" t="n">
        <v>%(U_L3L1)s</v>
      </c>
      <c r="P%(row)s" s="0" t="n">
        <f aca="false">E%(row)s*F%(row)s</f>
      </c>
      <c r="Q%(row)s" s="0" t="n">
        <f aca="false">G%(row)s*H%(row)s</f>
      </c>
      <c r="R%(row)s" s="0" t="n">
        <f aca="false">I%(row)s*J%(row)s</f>
      </c>
      <c r="S%(row)s" s="10" t="n">
        <f aca="false">SUM(P%(row)s:R%(row)s)</f>
      </c>
      <c r="T%(row)s" s="2" t="n">
        <f aca="false">S%(row)s/1000</f>
      </c>
      <c r="U%(row)s" s="3" t="n">
        <f aca="false">T%(row)s/1000</f>
      </c>
      <c r="V%(row)s" s="3" t="n">
        <v>%(freq)s</v>
      </c>
      <c r="W%(row)s" s="0" t="n">
        <v>%(count)s</v>
      </c>
    </row>
"""
sheet1_template_xml_top="""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr filterMode="false">
    <pageSetUpPr fitToPage="false"/>
  </sheetPr>
  <dimension ref="A1:AA1048576"/>
  <sheetViews>
    <sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="90" zoomScaleNormal="90" zoomScalePageLayoutView="100" workbookViewId="0">
      <pane xSplit="0" ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/>
      <selection pane="bottomLeft" activeCell="W3" activeCellId="0" sqref="W3"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" zeroHeight="false" outlineLevelRow="0" outlineLevelCol="0"/>
  <cols>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="1" min="1" style="0" width="23.3"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="2" min="2" style="1" width="15.57"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="3" min="3" style="1" width="7"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="4" min="4" style="1" width="11.17"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="5" min="5" style="0" width="5.57"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="6" min="6" style="0" width="5"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="7" min="7" style="0" width="5.57"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="8" min="8" style="0" width="5"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="9" min="9" style="0" width="5.57"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="10" min="10" style="0" width="5"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="11" min="11" style="0" width="5.57"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="12" min="12" style="0" width="4.85"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="15" min="13" style="0" width="7.43"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="18" min="16" style="0" width="7"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="19" min="19" style="0" width="12.82"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="20" min="20" style="2" width="9.43"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="21" min="21" style="3" width="10"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="22" min="22" style="0" width="11.28"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="23" min="23" style="0" width="8.53"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="24" min="24" style="0" width="22.28"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="25" min="25" style="0" width="7"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="26" min="26" style="0" width="11.28"/>
    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="1025" min="27" style="0" width="8.53"/>
  </cols>
  <sheetData>
    <row r="1" s="4" customFormat="true" ht="15" hidden="false" customHeight="false" outlineLevel="0" collapsed="false">
      <c r="A1" s="4" t="s">
        <v>0</v>
      </c>
      <c r="B1" s="5" t="s">
        <v>1</v>
      </c>
      <c r="C1" s="5" t="s">
        <v>2</v>
      </c>
      <c r="D1" s="5" t="s">
        <v>3</v>
      </c>
      <c r="E1" s="4" t="s">
        <v>4</v>
      </c>
      <c r="F1" s="4" t="s">
        <v>5</v>
      </c>
      <c r="G1" s="4" t="s">
        <v>6</v>
      </c>
      <c r="H1" s="4" t="s">
        <v>7</v>
      </c>
      <c r="I1" s="4" t="s">
        <v>8</v>
      </c>
      <c r="J1" s="4" t="s">
        <v>9</v>
      </c>
      <c r="K1" s="4" t="s">
        <v>10</v>
      </c>
      <c r="L1" s="4" t="s">
        <v>11</v>
      </c>
      <c r="M1" s="4" t="s">
        <v>12</v>
      </c>
      <c r="N1" s="4" t="s">
        <v>13</v>
      </c>
      <c r="O1" s="4" t="s">
        <v>14</v>
      </c>
      <c r="P1" s="4" t="s">
        <v>15</v>
      </c>
      <c r="Q1" s="4" t="s">
        <v>16</v>
      </c>
      <c r="R1" s="4" t="s">
        <v>17</v>
      </c>
      <c r="S1" s="4" t="s">
        <v>18</v>
      </c>
      <c r="T1" s="6" t="s">
        <v>19</v>
      </c>
      <c r="U1" s="7" t="s">
        <v>20</v>
      </c>
      <c r="V1" s="4" t="s">
        <v>21</v>
      </c>
      <c r="W1" s="4" t="s">
        <v>22</v>
      </c>
    </row>
"""
