import re
import zipfile
from datetime import datetime, timezone

template = {
    '_rels/.rels' : '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
</Relationships>
''',
    'docProps/app.xml' : '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Microsoft Excel</Application>
    <DocSecurity>0</DocSecurity>
    <ScaleCrop>false</ScaleCrop>
    <HeadingPairs>
        <vt:vector size="2" baseType="variant">
            <vt:variant>
                <vt:lpstr>Worksheets</vt:lpstr>
            </vt:variant>
            <vt:variant>
                <vt:i4>1</vt:i4>
            </vt:variant>
        </vt:vector>
    </HeadingPairs>
    <TitlesOfParts>
        <vt:vector size="1" baseType="lpstr">
            <vt:lpstr>Sheet1</vt:lpstr>
        </vt:vector>
    </TitlesOfParts>
    <Company></Company>
    <LinksUpToDate>false</LinksUpToDate>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>12.0000</AppVersion>
</Properties>
''',
    'docProps/core.xml' :'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator></dc:creator>
    <cp:lastModifiedBy></cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">_datetime_</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">_datetime_</dcterms:modified>
</cp:coreProperties>
''',
    'xl/_rels/workbook.xml.rels':'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
</Relationships>
''',
    'xl/worksheets/sheet1.xml' :'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <dimension ref="_dimension_" />
    <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0" />
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15" />
    _colWidth_ _sheetData_
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
</worksheet>
''',
    'xl/sharedStrings.xml' :'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="_sstCnt_" uniqueCount="_sstUCnt_">
    _sharedStr_
</sst>
''',
    'xl/styles.xml' :'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font>
            <sz val="11" />
            <name val="Calibri" />
            <family val="2" />
            <scheme val="minor" />
        </font>
    </fonts>
    <fills count="5">
        <fill>
            <patternFill patternType="none" />
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFC0FFC0" />
            </patternFill>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFFFFFC0" />
            </patternFill>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFFFC0C0" />
            </patternFill>
        </fill>
    </fills>
    <borders count="2">
        <border>
            <left /> <right /> <top /> <bottom />
            <diagonal />
        </border>
        <border>
            <left style="thin"> <color auto="1" /> </left>
            <right style="thin"> <color auto="1" /> </right>
            <top style="thin"> <color auto="1" /> </top>
            <bottom style="thin"> <color auto="1" /> </bottom>
            <diagonal />
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
    </cellStyleXfs>
    <cellXfs count="5">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
            <alignment horizontal="center" />
        </xf>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
            <alignment horizontal="center" />
        </xf>
        <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
            <alignment horizontal="center" />
        </xf>
        <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
            <alignment horizontal="center" />
        </xf>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0" />
    </cellStyles>
    <dxfs count="0" />
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />
</styleSheet>
''',
    'xl/workbook.xml' :'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505" />
    <bookViews>
        <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" />
    </bookViews>
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1" />
    </sheets>
    <calcPr calcId="124519" fullCalcOnLoad="1" />
</workbook>
''',
    '[Content_Types].xml' :'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Default Extension="xml" ContentType="application/xml" />
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
</Types>
'''
}
class MkXlsx():
    def __init__(self, fname):
        self.filename = fname
        self.num_rows = 1
        self.num_columns = 1
        self.rows = []
        self.strings = []
        self.col_width = []

    def _add_string(self, text):
        if text in self.strings:
            return self.strings.index(text)
        self.strings.append(text)
        return len(self.strings) - 1

    def _trim_split(self, str, style):
        s = re.split(r',\s*(?![^"]*"(?:[^"]*"[^"]*")*[^"]*$)', str)
        l = []
        for ss in s:
            l.append([ss.strip(' "'), style])
        return l

    def _cellname(self, r, c):
        return f'{chr(ord("A")+c-1)}{r}'

    def _make_dic(self):
        dic = {
            "_datetime_":  datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "_dimension_": f'A1:{self._cellname(self.num_rows, self.num_columns)}',
            "_sheetData_": self._make_sheet(),
            "_sharedStr_": self._make_sharedstr(),
            "_sstCnt_":    str(self.sstcnt),
            "_sstUCnt_":   str(len(self.strings)),
            "_colWidth_":  self._make_colwidth()
        }
        return dic

    def _make_sheet(self):
        sh = ''
        r = 1
        self.sstcnt = 0
        for row in self.rows:
            sh += f'<row r="{r}" spans="1:{self.num_columns}">'
            c = 1
            for cell in row:
                if cell != []:
                    s = cell['s']
                    v = cell['v']
                    if type(v) is str:
                        if len(v) > 0 and v[0] == '=':
                            sh += f'<c r="{self._cellname(r,c)}" s="{s}"><f>{v[1:]}</f><v></v></c>'
                        else:
                            v = self._add_string(v)
                            self.sstcnt += 1
                            sh += f'<c r="{self._cellname(r, c)}" s="{s}" t="s"><v>{v}</v></c>'
                    else:
                        sh += f'<c r="{self._cellname(r, c)}" s="{s}" t="n"><v>{v}</v></c>'
                else:
                    sh += f'<c r="{self._cellname(r, c)}" t="s"><v></v></c>'
                c += 1
            sh += f'</row>'
            r += 1
        if len(sh) > 0:
            return f'<sheetData>{sh}</sheetData>'
        return '<sheetData/>'
    
    def _make_sharedstr(self):
        sh = ''
        for s in self.strings:
            sh += f'<si><t>{str(s)}</t></si>'
        return sh

    def _make_colwidth(self):
        cols = ''
        for c in range(0, len(self.col_width)):
            if self.col_width[c] >= 0:
                cols += f'<col min="{c+1}" max="{c+1}" width="{self.col_width[c]}" customWidth="1" />'
        if len(cols)>0:
            return f'<cols>{cols}</cols>'
        return ''

    def _replace_dic(self, txt, dic):
        for item in dic:
            txt = re.sub(item, dic[item], txt)
        return txt

    def write(self, row, col, val, style=1):
        style = max(min(style, 4), 0)
        while row >= len(self.rows):
            self.rows.append([])
            self.num_rows = max(self.num_rows, row + 1)
        while col >= len(self.rows[row]):
            self.rows[row].append([])
            self.num_columns = max(self.num_columns, col + 1)
        self.rows[row][col] = {'v':val, 's':style}

    def set_column(self, col, width):
        while len(self.col_width) <= col:
            self.col_width.append(-1)
        self.col_width[col] = width

    def close(self):
        with zipfile.ZipFile(self.filename, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            dic = self._make_dic()
            zf.mkdir('_rels')
            zf.writestr('_rels/.rels', template['_rels/.rels'])
            zf.mkdir('docProps')
            zf.writestr('docProps/app.xml', template['docProps/app.xml'])
            zf.writestr('docProps/core.xml', self._replace_dic(template['docProps/core.xml'], dic))
            zf.mkdir('xl')
            zf.mkdir('xl/_rels')
            zf.writestr('xl/_rels/workbook.xml.rels', template['xl/_rels/workbook.xml.rels'])
            zf.mkdir('xl/worksheets')
            zf.writestr('xl/worksheets/sheet1.xml', self._replace_dic(template['xl/worksheets/sheet1.xml'], dic))
            zf.writestr('xl/sharedStrings.xml', self._replace_dic(template['xl/sharedStrings.xml'], dic))
            zf.writestr('xl/styles.xml', template['xl/styles.xml'])
            zf.writestr('xl/workbook.xml', template['xl/workbook.xml'])
            zf.writestr('[Content_Types].xml', template['[Content_Types].xml'])

if __name__ == "__main__":
    mkxlsx = MkXlsx('test.xlsx')
    mkxlsx.write(0, 0,"TEST")
    mkxlsx.write(1, 0, "Column-A", 2)
    mkxlsx.write(1, 1, "Column-B", 3)
    mkxlsx.write(1, 2, "Column-C", 4)
    mkxlsx.write(2, 0, "AAA")
    mkxlsx.write(2, 1, "BBB")
    mkxlsx.write(2, 2, "CCC")
    mkxlsx.set_column(0, 10)
    mkxlsx.set_column(1, 30)
    mkxlsx.close()