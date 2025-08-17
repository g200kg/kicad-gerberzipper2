#!/usr/bin/env python3

import sys
import os
import re
import json
import wx
import wx.grid
import wx.lib.scrolledpanel
import locale
import zipfile
import subprocess
import glob
import atexit
import shutil
import codecs
import inspect
import traceback
from kipy import KiCad
from kipy.proto.board.board_types_pb2 import BoardLayer
from kipy.proto.common.types import base_types_pb2, DocumentType, DocumentSpecifier

version = "0.5.0"
kicadCliCmd = "kicad-cli"

chsize = (10,20)
strtab = {}
execDialog = None
report = ''


layer_list = [
    {'name':'F.Cu',      'fname':'F_Cu',            'id':'F_Cu',        'fnamekey':'${filename(F.Cu)}'},
    {'name':'B.Cu',      'fname':'B_Cu',            'id':'B_Cu',        'fnamekey':'${filename(B.Cu)}'},
    {'name':'F.Adhes',   'fname':'F_Adhesive',      'id':'F_Adhes',     'fnamekey':'${filename(F.Adhes)}'},
    {'name':'B.Adhes',   'fname':'B_Adhesive',      'id':'B_Adhes',     'fnamekey':'${filename(B.Adhes)}'},
    {'name':'F.Paste',   'fname':'F_Paste',         'id':'F_Paste',     'fnamekey':'${filename(F.Paste)}'},
    {'name':'B.Paste',   'fname':'B_Paste',         'id':'B_Paste',     'fnamekey':'${filename(B.Paste)}'},
    {'name':'F.SilkS',   'fname':'F_Silkscreen',    'id':'F_SilkS',     'fnamekey':'${filename(F.SilkS)}'},
    {'name':'B.SilkS',   'fname':'B_Silkscreen',    'id':'B_SilkS',     'fnamekey':'${filename(B.SilkS)}'},
    {'name':'F.Mask',    'fname':'F_Mask',          'id':'F_Mask',      'fnamekey':'${filename(F.Mask)}'},
    {'name':'B.Mask',    'fname':'B_Mask',          'id':'B_Mask',      'fnamekey':'${filename(B.Mask)}'},
    {'name':'Dwgs.User', 'fname':'User_Drawings',   'id':'Dwgs_User',   'fnamekey':'${filename(Dwgs.User)}'},
    {'name':'Cmts.User', 'fname':'User_Comments',   'id':'Cmts_User',   'fnamekey':'${filename(Cmts.User)}'},
    {'name':'Eco1.User', 'fname':'User_Eco1',       'id':'Eco1_User',   'fnamekey':'${filename(Eco1.User)}'},
    {'name':'Eco2.User', 'fname':'User_Eco2',       'id':'Eco2_User',   'fnamekey':'${filename(Eco2.User)}'},
    {'name':'Edge.Cuts', 'fname':'Edge_Cuts',       'id':'Edge_Cuts',   'fnamekey':'${filename(Edge.Cuts)}'},
    {'name':'F.CrtYd',   'fname':'F_Courtyard',     'id':'F_CrtYd',     'fnamekey':'${filename(F.CrtYd)}'},
    {'name':'B.CrtYd',   'fname':'B_Courtyard',     'id':'B_CrtYd',     'fnamekey':'${filename(B.CrtYd)}'},
    {'name':'F.Fab',     'fname':'F_Fab',           'id':'F_Fab',       'fnamekey':'${filename(F.Fab)}'},
    {'name':'B.Fab',     'fname':'B_Fab',           'id':'B_Fab',       'fnamekey':'${filename(B.Fab)}'},
    {'name':'In1.Cu',    'fname':'In1_Cu',          'id':'In1_Cu',      'fnamekey':'${filename(In1.Cu)}'},
    {'name':'In2.Cu',    'fname':'In2_Cu',          'id':'In2_Cu',      'fnamekey':'${filename(In2.Cu)}'},
    {'name':'In3.Cu',    'fname':'In3_Cu',          'id':'In3_Cu',      'fnamekey':'${filename(In3.Cu)}'},
    {'name':'In4.Cu',    'fname':'In4_Cu',          'id':'In4_Cu',      'fnamekey':'${filename(In4.Cu)}'},
    {'name':'In5.Cu',    'fname':'In5_Cu',          'id':'In5_Cu',      'fnamekey':'${filename(In5.Cu)}'},
    {'name':'In6.Cu',    'fname':'In6_Cu',          'id':'In6_Cu',      'fnamekey':'${filename(In6.Cu)}'}
]

default_settings = {
  "Name":"ManufacturersName",
  "Description":"description",
  "URL":"https://example.com/",
  "GerberDir":"Gerber",
  "ZipFilename":"*.ZIP",
  "Layers": {
    "F.Cu":"",
    "B.Cu":"",
    "F.Paste":"",
    "B.Paste":"",
    "F.SilkS":"",
    "B.SilkS":"",
    "F.Mask":"",
    "B.Mask":"",
    "Edge.Cuts":"",
    "In1.Cu":"",
    "In2.Cu":"",
    "In3.Cu":"",
    "In4.Cu":"",
    "In5.Cu":"",
    "In6.Cu":""
  },
  "PlotBorderAndTitle":False,
  "PlotFootprintValues":True,
  "PlotFootprintReferences":True,
  "UseAuxOrigin":False,
  "CoordinateFormat46":True,
  "SubtractMaskFromSilk":True,
  "UseExtendedX2format": False,
  "IncludeNetlistInfo":False,
  "Drill": {
    "Drill":"",
    "DrillMap":"",
    "NPTH":"",
    "NPTHMap":"",
    "Report":""
  },
  "DrillUnitMM":True,
  "MirrorYAxis":False,
  "MinimalHeader":False,
  "MergePTHandNPTH":False,
  "RouteModeForOvalHoles":True,
  "ZerosFormat":{
    "DecimalFormat":True,
    "SuppressLeading":False,
    "SuppressTrailing":False,
    "KeepZeros":False
  },
  "MapFileFormat":{
    "PostScript":False,
    "Gerber":False,
    "DXF":False,
    "SVG":False,
    "PDF":True
  },
  "OptionalFiles":[],
  "BomFile":{
    "TopFilename":"*-BOM-Top.csv",
    "BottomFilename":"*-BOM-Bottom.csv",
    "MergeSide":False,
    "IncludeTHT":False,
    "Header":"Comment, Designator, Footprint, Part#, Qty",
    "Row": "\"${val}\",\"${ref}\",\"${fp}\",\"${PN}\",${qty}",
    "Footer":""
  },
  "PosFile":{
    "TopFilename":"*-POS-Top.csv",
    "BottomFilename":"*-POS-Bottom.csv",
    "MergeSide":False,
    "IncludeTHT":False,
    "Header": "Designator, PosX, PosY, Side, Rotation, Package, Type",
    "Row": "\"${ref}\",${x},${y},\"${side}\",${rot},\"${fp}\",\"${type}\"",
    "Footer":""
  }
}

class ExecDialog(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, 'GerberZipper 2 Exec')
        self.SetClientSize(Em(60,20))
        panel = wx.Panel(self)
        self.text = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, pos=Em(1,1),size=Em(58,16))
#        self.text.Disable()
        self.button = wx.Button(panel, wx.ID_OK, "Close", pos=Em(5,18),size=Em(10,1.5))
        self.button.Hide()
        self.button.Bind(wx.EVT_BUTTON,self.Close)
        self.Show()
    def Progress(self, s):
        self.text.SetValue(s)
        self.Refresh()
        self.Update()
    def Close(self, e):
        print('ExecDialog.Close')
        self.Destroy()
        e.Skip()
    def Complete(self, stat):
        self.button.Show()
        if not stat:
            alert('Complete')

def ExecProgress(self, s, icon=0):
    if self.FindWindowByName('GerberZipper 2 Exec') is None:
        return ExecDialog()
    else:
        return None

def alert(s, icon=0):
    dialog = wx.MessageDialog(None, s, 'Gerber Zipper 2', icon)
    r = dialog.ShowModal()
    return r

def InitEm():
    global chsize
    dc = wx.ScreenDC()
    font = wx.Font(pointSize=10,family=wx.DEFAULT,style=wx.NORMAL,weight=wx.NORMAL)
    dc.SetFont(font)
    tx = dc.GetTextExtent("M")
    chsize = (tx[0],tx[1]*1.5)

def Em(x,y,dx=0,dy=0):
    return (int(chsize[0]*x+dx), int(chsize[1]*y+dy))

def getindex(s):
    for i in range(len(layer_list)):
        if layer_list[i]['name']==s:
            return i
    return -1

def getid(s):
    for i in range(len(layer_list)):
        if layer_list[i]['name']==s:
            return layer_list[i]['id']
    return None

def getfname(s):
    for i in range(len(layer_list)):
        if layer_list[i]['name']==s:
            return layer_list[i]['fname']
    return None

def getstr(s):
#    lang = lang if lang else 'default'
#    lang = locale.getdefaultlocale()[0]
    lang = 'default'
    tab = strtab['default']
    if (lang in strtab):
        tab = strtab[lang]
    else:
        for x in strtab:
            if lang[0:3] in x:
                tab = strtab[x]
    return tab.get(s, strtab['default'].get(s, s))

def forcedel(fname):
    if os.path.exists(fname):
        os.remove(fname)

def forceren(src, dst):
    if(src==dst):
        return
    forcedel(dst)
    if os.path.exists(src):
        os.rename(src, dst)

def getsubkey(s):
    l = s.split(' ')
    subkeys = {}
    for sk in l:
        sks = sk.split(':')
        if(len(sks) == 2):
            subkeys[sks[0]] = sks[1]
    return subkeys

def tabexp(str,tabTable):
    strList = str.split('\t')
    result = ''
    curColumn = 0
    strIdx = 0
    nextTab = 0
    for strCur in strList:
        strCur = strCur[:(tabTable[strIdx]-1)]
        while curColumn < nextTab:
            curColumn += 1
            result += ' '
        result += strCur
        curColumn += len(strCur)
        nextTab += tabTable[strIdx]
        strIdx += 1
    return result

def strreplace(s,d):
    m = re.findall('\${([0-9a-zA-Z|_]*)}',s)
    r = {}
    v = {}
    for i in m:
        r[i] = i.split('|')
        for ii in r[i]:
            if(ii in d):
                v[i] = d[ii]
                break
            else:
                v[i] = ''
    for k in v:
        s = s.replace('${'+k+'}', str(v[k]))
    return s

def isNum(x):
    try:
        float(x)
    except ValueError:
        return False
    else:
        return True

def renamefile(olddir, oldfname, newdir, newfname, zipfile):
    if oldfname == '' or newfname == '':
        return
    oldf = os.path.join(olddir, oldfname)
    newf = os.path.join(newdir, newfname)
    if os.path.exists(oldf):
#        print(f'{oldf} => {newf}')
        if os.path.exists(newf):
            os.remove(newf)
        os.rename(oldf, newf)
        if zipfile != None:
            zipfile.append(newf)

class tableFile():
    def __init__(self, fn):
        self.row = 0
        self.fname = fn
        self.tabs = []
        self.xlsxReady = 1
        if fn.endswith('.csv'):
            self.type = 'csv'
            try:
                self.f = open(fn, mode='w', encoding='utf-8')
            except Exception as err:
                alert('File creation error \n\n File : %s\n' % (fn), wx.ICON_ERROR)
                raise
        elif fn.endswith('.xlsx'):
            self.xlsxReady = 1
            try:
                import xlsxwriter
            except ModuleNotFoundError:
                self.xlsxReady = 0
            self.type = 'xlsx'
            if self.xlsxReady == 0:
                self.type = 'csv'
                self.fname = self.fname[0:-5] + '.csv'
                try:
                    self.f = open(self.fname, mode='w', encoding='utf-8')
                except Exception as err:
                    alert('File creation eror \n\n File : %s\n' % (fn), wx.ICON_ERROR)
                    raise
            else:
                try:
                    self.f = open(self.fname, mode='w', encoding='utf-8')
                except Exception as err:
                    alert('File creation eror \n\n File : %s\n' % (fn), wx.ICON_ERROR)
                    raise
                self.f.close()
                self.xlsx = xlsxwriter.Workbook(self.fname)
                self.sheet = self.xlsx.add_worksheet('Sheet1')
                self.HeaderFormat = self.xlsx.add_format()
                self.HeaderFormat.set_bold()
                self.HeaderFormat.set_bg_color('yellow')
                self.HeaderFormat.set_align('center')
                self.HeaderFormat.set_border(1)
                self.BodyFormat = self.xlsx.add_format()
                self.BodyFormat.set_text_wrap()
                self.BodyFormat.set_border(1)
                self.BodyFormat.set_align('center')
        else:
            self.type = 'txt'
            self.f = open(fn, mode='w', encoding='utf-8')
        self.ini=0
    
    def setTabs(self, tabs):
        self.tabs = tabs
        if self.type == 'xlsx':
            for i in range(len(self.tabs)):
                self.sheet.set_column(i, i, float(self.tabs[i]))

    def deleteSubkeys(self, str):
        s = str.strip('"').split(' ')
        s2 = ''
        for ss in s:
            if ':' not in ss:
                if s2 != '':
                    s2 += ' '
                s2 += ss
        return s2

    def addLine(self, line, dic, format):
        if line == None:
            return
        if self.type == 'xlsx':
            cells = line.split(',')
            res = []
            for cell in cells:
                res.append(strreplace(cell, dic))
            if format == 'Header':
                font = self.HeaderFormat
            else:
                font = self.BodyFormat
            col = 0
            for cell in res:
                if isNum(cell[0]):
                    self.sheet.write(self.row, col, float(cell), font)
                else:
                    s = self.deleteSubkeys(cell)
                    self.sheet.write(self.row, col, s, font)
                col += 1
        elif self.type == 'csv':
            rstr=strreplace(line,dic)
            self.f.write(rstr+'\n')
        else:
            if format == 'Header':
                self.f.write(line + '\n')
            else:
                cells = line.split('\t')
                res = []
                tabcnt = 0
                for cell in cells:
                    r = strreplace(cell, dic)
                    r2 = r
                    if(self.tabs):
                        tablen = self.tabs[tabcnt]
                        r2 = (r + ' ' * tablen)[0:tablen-1]
                    res.append(r2)
                    tabcnt += 1
                self.f.write(' '.join(res) + '\n')
        self.row += 1

    def close(self):
        if self.type == 'xlsx':
            self.xlsx.close()
        else:
            self.f.close()

def message(s):
    print('GerberZipper: '+s)

def FpToDict(fp):
    val = fp.value_field.text.value
    ref = fp.reference_field.text.value
    pak = fp.definition.id.name
    layer = 0 if fp.layer == BoardLayer.BL_F_Cu else 1
    side = ['Top', 'Bottom'][layer]
    side1 = side[0]
    mount = ['','Through hole','SMD','Unspecified'][fp.attributes.mounting_style]
    return {'val':val, 'ref':ref, 'pak':pack, 'layer':layer, 'side':side, 'side1':side[0], 'type':mount}

def GetBoard():
    global kicad
    try:
        kicad = KiCad()
        print(f"KiCad version: {kicad.get_version().full_version}")
    except BaseException as e:
        alert(f"Not connected to KiCad: {e}")
        exit()
    
    print(f"Api version {kicad.get_api_version()}")
    try:
        if kicad.check_version():
            print("KiCad version and kicad-python version match :)")
    except BaseException as e:
        print(f"{e}")
    try:
        board = kicad.get_board()
    except BaseException as e:
        alert(f"PCB Open Error : ({e})")
        exit()
    return board

def GerberExec(self, board):
    global report
    self.settings = self.Get()
    zipfiles = []
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags = subprocess.CREATE_NEW_CONSOLE | subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE

    # Gerber
    opt = ''
    layers = self.settings['Layers']
    lystr = ''
    kylist = layers.keys()
    print(layers)
    for k in kylist:
        if layers[k] != '':
            lystr += k
            lystr += ','
    lystr = lystr[:-1]
    if self.settings['PlotBorderAndTitle']:
        opt += '--ibt '
    if not self.settings['PlotFootprintValues']:
        opt += '--ev '
    if not self.settings['PlotFootprintReferences']:
        opt += '--erd '
    if self.settings['SubtractMaskFromSilk']:
        opt += '--subtract-soldermask '
    if not self.settings['UseExtendedX2format']:
        opt += '--no-x2 '
    if not self.settings['IncludeNetlistInfo']:
        opt += '--no-netlist '
    if not self.settings['CoordinateFormat46']:
        opt += '--precision 5 '
    if self.settings['UseAuxOrigin']:
        opt += '--use-drill-file-origin '
    cmd = f'kicad-cli pcb export gerbers {opt} --no-protel-ext -o {self.temp_dir} -l {lystr} {self.board_path}'
    print(f'cmd : {cmd}')
    ret = 0
    try:
        ret = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo).stdout
    except Exception as err:
        alert(f'error \n\n {err}')
    for k in kylist:
        if layers[k] != '':
            fname = getfname(k)
            renamefile(self.temp_dir, f'$$$-{fname}.gbr', self.gerber_dir, f'{layers[k].replace("*",self.basename)}', zipfiles)

    # Drill
    opt = ''
    mext = 'pdf'
    mapf = 'PDF'
    if self.settings['MirrorYAxis']:
        opt += '--excellon-mirror-y '
    if self.settings['MinimalHeader']:
        opt += '--excellon-min-header '
    if not self.settings['MergePTHandNPTH']:
        opt += '--excellon-separate-th '
    if self.settings['RouteModeForOvalHoles']:
        opt += '--excellon-oval-format route '
    if not self.settings['DrillUnitMM']:
        opt += '--u in '
    if self.settings['UseAuxOrigin']:
        opt += '--drill-origin plot '
    zerof = [k for k,v in self.settings['ZerosFormat'].items() if v == True][0]
    if zerof != 'DecimalFormat':
        opt += '--excellon-zeros-format '
        opt += {'SuppressLeading':'suppressleading ', 'SuppressTrailing':'suppresstrailing ', 'KeepZeros':'keep '}[zerof]
    if self.settings['Drill']['DrillMap'] != '':
        opt += '--generate-map '
        mapf = [k for k,v in self.settings['MapFileFormat'].items() if v == True][0]
        opt += '--map-format '
        opt += {'PostScript':'ps ','Gerber':'gerberx2 ','DXF':'dxf ','SVG':'svg ','PDF':'pdf '}[mapf]
        mext = {'PostScript':'ps','Gerber':'gbr','DXF':'dxf','SVG':'svg','PDF':'pdf'}[mapf]
    cmd = f'kicad-cli pcb export drill {opt} -o {self.temp_dir} --format excellon {self.board_path}'
    print(f'cmd : {cmd}')
    ret = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo).stdout
    if self.settings['MergePTHandNPTH']:
        renamefile(self.temp_dir, '$$$.drl', self.gerber_dir, self.settings['Drill']['Drill'].replace('*', self.basename), zipfiles)
        if self.settings['Drill']['DrillMap'] != '':
            renamefile(self.temp_dir, f'$$$-drl_map.{mext}', self.gerber_dir, self.settings['Drill']['DrillMap'].replace('*', self.basename), zipfiles)
    else:
        renamefile(self.temp_dir, '$$$-PTH.drl', self.gerber_dir, self.settings['Drill']['Drill'].replace('*', self.basename), zipfiles)
        renamefile(self.temp_dir, '$$$-NPTH.drl', self.gerber_dir, self.settings['Drill']['NPTH'].replace('*', self.basename), zipfiles)
        if self.settings['Drill']['DrillMap'] != '':
            renamefile(self.temp_dir, f'$$$-PTH-drl_map.{mext}', self.gerber_dir, self.settings['Drill']['DrillMap'].replace('*', self.basename), zipfiles)
        if self.settings['Drill']['NPTHMap'] != '':
            renamefile(self.temp_dir, f'$$$-NPTH-drl_map.{mext}', self.gerber_dir, self.settings['Drill']['NPTHMap'].replace('*', self.basename), zipfiles)

    # Zip
    zipfname = f'{self.gerber_dir}/{self.zipfilename.GetValue().replace("*",self.basename)}'
    with zipfile.ZipFile(zipfname,'w',compression=zipfile.ZIP_DEFLATED) as f:
        for i in range(len(zipfiles)):
            fnam = zipfiles[i]
            if os.path.exists(fnam):
                f.write(fnam, os.path.basename(fnam))
    report += (getstr('COMPLETE') % zipfname)+'\n'

def RefillExec(self, board):
    global report
    print('RefillExec')
    r = board.refill_zones(block=True)
    print(r)
    report += 'Done\n'

def DrcExec(self, board):
    global report
    print('DrcExec')
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags = subprocess.CREATE_NEW_CONSOLE | subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    opt = ''
    cmd = f'kicad-cli pcb drc {opt} -o {self.gerber_dir}/{self.basename}.rpt --format report {self.board_path}'
    print(f'cmd : {cmd}')
    ret = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo).stdout
    report += ret

def FabExec(self, board):
    global report
    print('FabExec')
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags = subprocess.CREATE_NEW_CONSOLE | subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    opt = ''
    cmd = f'kicad-cli pcb export pdf {opt} -o {self.gerber_dir}/{self.basename}.pdf -l Edge.Cuts,F.Fab {self.board_path}'
    print(f'cmd : {cmd}')
    ret = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo).stdout
    report += ret

def BomPosExec(self, board):
    global report
    self.settings = self.Get()
    board_basename = (os.path.splitext(os.path.basename(board.name)))[0]

    # BOM
    bomParam = self.settings.get('BomFile',{})
    bom_fnameT = ''
    bom_fnameB = ''
    fnameT = bomParam.get('TopFilename','')
    fnameB = bomParam.get('BottomFilename','')
    if len(fnameT)>0:
        bom_fnameT = f'{self.gerber_dir}/{fnameT.replace("*", self.basename)}'
    if len(fnameB)>0 and not bomParam.get('MergeSide'):
        bom_fnameB = f'{self.gerber_dir}/{fnameB.replace("*", self.basename)}'
    print('Top.BOM:'+bom_fnameT+'\nBottom.BOM:'+bom_fnameB)
    bomList = [{},{}]
    for fp in board.get_footprints():
        val = fp.value_field.text.value
        ref = fp.reference_field.text.value
        pak = fp.definition.id.name
        layer = 0 if fp.layer == BoardLayer.BL_F_Cu else 1
        side = ['Top', 'Bottom'][layer]
        side1 = side[0]
        mount = ['', 'Through hole', 'SMD', 'Unspecified'][fp.attributes.mounting_style]
        dnp = fp.attributes.do_not_populate
        exbom = fp.attributes.exclude_from_bill_of_materials
        expos = fp.attributes.exclude_from_position_files
        if mount == 'SMD' or (mount == 'Through hole' and bomParam.get('IncludeTHT')):
            if bomParam.get('MergeSide'):
                layer = 0
            if val in bomList[layer]:
                bomList[layer][val]['ref'] += ',' + ref
                bomList[layer][val]['qty'] += 1
            else:
                bomList[layer][val] = {'val':val, 'ref':ref, 'fp':pak, 'qty':1, 'type':mount, 'side':side, 'side1':side1}
                bomList[layer][val].update(getsubkey(val))
    rowformat = bomParam.get('Row')
    header = bomParam.get('Header','')
    if len(bom_fnameT)>0:
        tfBomTop = tableFile(bom_fnameT)
        tfBomTop.setTabs(bomParam.get('Tabs'))
        if header != "":
            tfBomTop.addLine(header, {}, 'Header')
        itemNum = 1
        for val in bomList[0]:
            bomList[0][val]['num'] = itemNum
            tfBomTop.addLine(rowformat, bomList[0][val], 'Body')
            itemNum += 1
        tfBomTop.close()
    if len(bom_fnameB)>0:
        tfBomBottom = tableFile(bom_fnameB)
        tfBomBottom.setTabs(bomParam.get('Tabs'))
        if header != "":
            tfBomBottom.addLine(header, {}, 'Header')
        itemNum = 1
        for val in bomList[1]:
            bomList[1]['num'] = itemNum
            tfBomBottom.addLine(rowformat, bomList[1][val], 'Body')
            itemNum += 1
        tfBomBottom.close()

    # POS
    posParam = self.settings.get('PosFile',{})
    fnameT = posParam.get('TopFilename','')
    fnameB = posParam.get('BottomFilename','')
    pos_fnameT = ''
    pos_fnameB = ''
    if len(fnameT)>0:
        pos_fnameT = f'{self.gerber_dir}/{fnameT.replace("*", board_basename)}'
    if len(fnameB)>0 and not posParam.get('MergeSide'):
        pos_fnameB = f'{self.gerber_dir}/{fnameB.replace("*", board_basename)}'
    try:
        if len(pos_fnameT)>0:
            tfPosTop = tableFile(pos_fnameT)
            tfPosTop.setTabs(posParam.get('Tabs'))
            tfPosTop.addLine(posParam.get('Header'), {'side':'Top'}, 'Header')
        if len(pos_fnameB)>0:
            tfPosBottom = tableFile(pos_fnameB)
            tfPosBottom.setTabs(posParam.get('Tabs'))
            tfPosBottom.addLine(posParam.get('Header'), {'side':'Bottom'}, 'Header')
        print('Top.POS:'+pos_fnameT+'\nBottom.POS:'+pos_fnameB)
        origin = board.get_origin(2)    # drill-origin
        if not self.settings['UseAuxOrigin']:
            origin.x = 0
            origin.y = 0
        print(f'origin : {origin}')
        rowformat = posParam.get('Row')
        for fp in board.get_footprints():
            val = fp.value_field.text.value
            ref = fp.reference_field.text.value
            pos = fp.position
            rot = fp.orientation.degrees
            pak = fp.definition.id.name
            layer = 0 if fp.layer == BoardLayer.BL_F_Cu else 1
            side = ['Top', 'Bottom'][layer]
            side1 = side[0]
            mount = ['', 'Through hole', 'SMD', 'Unspecified'][fp.attributes.mounting_style]
            dnp = fp.attributes.do_not_populate
            exbom = fp.attributes.exclude_from_bill_of_materials
            expos = fp.attributes.exclude_from_position_files
            print(f'FP : {ref} : {pos} {rot} {side} {mount} {pak}')
            if mount == 'SMD' or (mount == 'Through hole' and posParam.get('IncludeTHT')):
                dict = {'val':val, 'ref':ref, 'x':pos.x-origin.x, 'y':pos.y-origin.y, 'fp':pak, 'type':mount, 'side':side, 'side1':side1, 'rot':rot}
#                        dict.update(subkey)
                row = strreplace(rowformat, dict)
                tabs = posParam.get('Tabs')
                if tabs:
                    row = tabexp(row, tabs)
                if side1 == 'T' or posParam.get('MergeSide'):
                    if tfPosTop != None:
                        tfPosTop.addLine(rowformat, dict, 'Body')
                else:
                    if tfPosBottom != None:
                        tfPosBottom.addLine(rowformat, dict, 'Body')

        if tfPosTop != None:
            tfPosTop.addLine(posParam.get('Footer'), {'side':'top'}, 'Body')
            tfPosTop.close()
        if tfPosBottom != None:
            tfPosBottom.addLine(posParam.get('Footer'), {'side':'bottom'}, 'Body')
            tfPosBottom.close()
    except Exception as e:
        print(e)
        return
    ret = getstr('BOMPOSCOMPLETE') % (bom_fnameT, bom_fnameB, pos_fnameT, pos_fnameB)
    report += ret
#    alert(ret, wx.ICON_INFORMATION)

class GerberZipper2():
    def Run(self):
        class Dialog(wx.Dialog):
            def __init__(self, parent):
                global strtab
                atexit.register(self.ClosePlugin)
                prefix_path = os.path.join(os.path.dirname(__file__))
                self.pluginSettings = {}
                self.icon_file_name = os.path.join(os.path.dirname(__file__), 'Assets/icon48.png')
                self.manufacturers_dir = os.path.join(os.path.dirname(__file__), 'Manufacturers')
                manufacturers_list = glob.glob('%s/*.json' % self.manufacturers_dir)
                self.json_data = {}
                for fname in manufacturers_list:
                    try:
                        d = json.load(codecs.open(fname, 'r', 'utf-8'))
                        self.json_data[d['Name']] = json.load(codecs.open(fname, 'r', 'utf-8'))
                    except Exception as err:
                        tb = sys.exc_info()[2]
                        alert('JSON error \n\n File : %s\n%s' % (os.path.basename(fname),err.with_traceback(tb)), wx.ICON_WARNING)
                self.settings_fname = os.path.join(prefix_path, 'GerberZipper2.json')
                if os.path.exists(self.settings_fname):
                    self.pluginSettings = json.load(open(self.settings_fname))
                else:
                    self.pluginSettings["Recent"] = next(iter(self.json_data))
                    self.pluginSettings["RefillExec"] = True
                    self.pluginSettings["DrcExec"] = True
                    self.pluginSettings["FabExec"] = True
                    self.pluginSettings["BomPosExec"] = True
                self.locale_dir = os.path.join(os.path.dirname(__file__), "Locale")
                locale_list = glob.glob('%s/*.json' % self.locale_dir)
                strtab = {}
                for fpath in locale_list:
                    fname = os.path.splitext(os.path.basename(fpath))[0]
                    strtab[fname] = json.load(codecs.open(fpath, 'r', 'utf-8'))
                InitEm()
                self.szPanel = [Em(75,10), Em(75,38)]
                wx.Dialog.__init__(self, parent, id=-1, title=f'Gerber Zipper 2 (ver. {version})', style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
                self.panel = wx.Panel(self)
                self.SetIcon(wx.Icon(self.icon_file_name))

                manufacturers_arr=[]
                for item in self.json_data.keys():
                    manufacturers_arr.append(item)
                wx.StaticText(self.panel, -1, getstr('LABEL'), size=Em(70,1), pos=Em(1,1))
                wx.StaticText(self.panel, -1, getstr('MENUFACTURERS'),size=Em(14,1), pos=Em(1,2.5))

                self.manufacturers = wx.ComboBox(self.panel, -1, 'Select Manufacturers', size=Em(30,1.5), pos=Em(16,2.5), choices=manufacturers_arr, style=wx.CB_READONLY)
                wx.StaticText(self.panel, -1, getstr('URL'),size=Em(14,1), pos=Em(1,4))
                self.url = wx.TextCtrl(self.panel, -1, '', size=Em(30,1), pos=Em(16,4), style=wx.TE_READONLY)
                wx.StaticText(self.panel, -1, getstr('GERBERDIR'),size=Em(14,1), pos=Em(1,5.1))
                self.gerberdir = wx.TextCtrl(self.panel, -1, '',size=Em(30,1), pos=Em(16,5.1))
                wx.StaticText(self.panel, -1, getstr('ZIPFNAME'),size=Em(14,1), pos=Em(1,6.2))
                self.zipfilename = wx.TextCtrl(self.panel, -1, '',size=Em(30,1), pos=Em(16,6.2))
                wx.StaticText(self.panel, -1, getstr('DESCRIPTION'),size=Em(14,1), pos=Em(1,7.3))
                self.label = wx.StaticText(self.panel, -1, '',size=Em(45,1), pos=Em(16,7.3))

                self.opt_RefillExec = wx.CheckBox(self.panel, -1, getstr('REFILLEXEC'), size=Em(15,1),pos=Em(52,3))
                self.opt_DrcExec = wx.CheckBox(self.panel, -1, getstr('DRCEXEC'), size=Em(15,1),pos=Em(52,4))
                self.opt_FabExec = wx.CheckBox(self.panel, -1, getstr('FABEXEC'), size=Em(15,1),pos=Em(52,5))
                self.opt_BomPosExec = wx.CheckBox(self.panel, -1, getstr('BOMPOSEXEC'), size=Em(15,1),pos=Em(52,6))

                self.detailbtn = wx.ToggleButton(self.panel, -1, getstr('DETAIL'),size=Em(15,1),pos=Em(2,8.5))
                self.execbtn = wx.Button(self.panel, -1, getstr('EXEC'),size=Em(15,1),pos=Em(18,8.5))
                self.clsbtn = wx.Button(self.panel, -1, getstr('CLOSE'),size=Em(15,1),pos=Em(50,8.5))
                self.manufacturers.Bind(wx.EVT_COMBOBOX, self.OnManufacturers)
                self.detailbtn.Bind(wx.EVT_TOGGLEBUTTON, self.OnDetail)
                self.execbtn.Bind(wx.EVT_BUTTON, self.OnExec)
                self.clsbtn.Bind(wx.EVT_BUTTON, self.OnClose)

                self.panel2 = wx.lib.scrolledpanel.ScrolledPanel(self.panel, -1, pos=Em(0,10), size=Em(75,20), style=wx.BORDER_SUNKEN)
                self.panel2.SetScrollbars(20,20,50,50)
                
                wx.StaticText(self.panel2, -1, 'ZIP contents', pos=Em(1,0), size=Em(12,1))
                wx.StaticLine(self.panel2, -1, pos=Em(9,0.5), size=(Em(56,1)[0],2))
                wx.StaticBox(self.panel2, -1,'Gerber', pos=Em(2,1), size=Em(65,13))
                wx.StaticBox(self.panel2, -1,'Drill', pos=Em(2,14), size=Em(65,8))
                wx.StaticBox(self.panel2, -1,'Origin', pos=Em(2,22), size=Em(65,3))
                wx.StaticBox(self.panel2, -1,'Other', pos=Em(2,25), size=Em(65,3))

                self.layer = wx.grid.Grid(self.panel2, -1, size=Em(18,11), pos=Em(3,2))
                self.layer.SetColLabelSize(Em(1,1)[1])
                self.layer.DisableDragColSize()
                self.layer.DisableDragRowSize()
                self.layer.CreateGrid(len(layer_list), 2)
                self.layer.SetColLabelValue(0, 'Layer')
                self.layer.SetColLabelValue(1, 'Filename')
                self.layer.SetRowLabelSize(1)
                self.layer.ShowScrollbars(wx.SHOW_SB_NEVER,wx.SHOW_SB_DEFAULT)
                self.layer.SetColSize(0, Em(7,1)[0])
                self.layer.SetColSize(1, Em(9,1)[0])
                for i in range(len(layer_list)):
                    self.layer.SetCellValue(i, 0, layer_list[i]['name'])
                    self.layer.SetReadOnly(i, 0, isReadOnly=True)
                self.opt_PlotBorderAndTitle = wx.CheckBox(self.panel2, -1, 'PlotBorderAndTitle', pos=Em(23,2))
                self.opt_PlotFootprintValues = wx.CheckBox(self.panel2, -1, 'PlotFootprintValues', pos=Em(23,3))
                self.opt_PlotFootprintReferences = wx.CheckBox(self.panel2, -1, 'PlotFootprintReferences', pos=Em(23,4))
                self.opt_SubtractMaskFromSilk = wx.CheckBox(self.panel2, -1, 'SubtractMaskFromSilk', pos=Em(23, 5))
                self.opt_UseExtendedX2format = wx.CheckBox(self.panel2, -1, 'UseExtendedX2format', pos=Em(23, 6))
                self.opt_CoordinateFormat46 = wx.CheckBox(self.panel2, -1, 'CoordinateFormat46', pos=Em(23, 7))
                self.opt_IncludeNetlistInfo = wx.CheckBox(self.panel2, -1, 'IncludeNetlistInfo', pos=Em(23, 8))

                self.opt_UseAuxOrigin = wx.CheckBox(self.panel2, -1, 'UseAuxOrigin', pos=Em(23,23))

                self.drill = wx.grid.Grid(self.panel2, -1, size=Em(18,6,1,0), pos=Em(3,15))
                self.drill.DisableDragColSize()
                self.drill.DisableDragRowSize()
                self.drill.CreateGrid(5, 2)
                self.drill.DisableDragGridSize()
                self.drill.ShowScrollbars(wx.SHOW_SB_NEVER,wx.SHOW_SB_NEVER)
                self.drill.SetColLabelValue(0, 'Drill')
                self.drill.SetColLabelValue(1, 'Filename')
                self.drill.SetRowLabelSize(1)
                self.drill.SetColSize(0, Em(9,1)[0])
                self.drill.SetColSize(1, Em(9,1)[0])
                drillfile = ['Drill', 'DrillMap', 'NPTH', 'NPTHMap', 'Report']
                self.drill.SetColLabelSize(Em(1,1)[1])
                for i in range(len(drillfile)):
                    self.drill.SetCellValue(i, 0, drillfile[i])
                    self.drill.SetReadOnly(i, 0, True)
                    self.drill.SetRowSize(i, Em(1,1)[1])
                self.opt_MirrorYAxis = wx.CheckBox(self.panel2, -1, 'MirrorYAxis', pos=Em(23,16))
                self.opt_MinimalHeader = wx.CheckBox(self.panel2, -1, 'MinimalHeader', pos=Em(23,17))
                self.opt_MergePTHandNPTH = wx.CheckBox(self.panel2, -1, 'MergePTHandNPTH', pos=Em(23,18))
                self.opt_RouteModeForOvalHoles = wx.CheckBox(self.panel2, -1, 'RouteModeForOvalHoles', pos=Em(23,19))
                wx.StaticText(self.panel2, -1, 'Drill Unit :', pos=Em(43,16))
                self.opt_DrillUnit = wx.ComboBox(self.panel2, -1, '', choices=('inch','mm'), style=wx.CB_READONLY, pos=Em(54,16), size=Em(8,1))
                self.opt_DrillUnit.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)
                wx.StaticText(self.panel2, -1, 'Zeros :', pos=Em(43,17.5))
                self.opt_ZerosFormat = wx.ComboBox(self.panel2, -1, '', choices=('DecimalFormat','SuppressLeading','SuppresTrailing', 'KeepZeros'), pos=Em(54,17.5), size=Em(12,1), style=wx.CB_READONLY)
                self.opt_ZerosFormat.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)
                wx.StaticText(self.panel2, -1, 'MapFileFormat :', pos=Em(43,19))
                self.opt_MapFileFormat = wx.ComboBox(self.panel2, -1, '', choices=('PostScript','Gerber','DXF','SVG','PDF'), pos=Em(54,19), size=Em(8,1), style=wx.CB_READONLY)
                self.opt_MapFileFormat.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)

                self.opt_OptionalLabel = wx.StaticText(self.panel2, -1, 'OptionalFile:', pos=Em(4,26))
                self.opt_OptionalFile = wx.TextCtrl(self.panel2, -1, '', pos=Em(15,26), size=Em(12,1))
                self.opt_OptionalContent = wx.TextCtrl(self.panel2, -1, '', pos=Em(28,26), size=Em(37,1))

                wx.StaticText(self.panel2, -1, 'BOM/POS', size=Em(7,1), pos=Em(1, 29))
                wx.StaticLine(self.panel2, -1, size=(Em(60,1)[0],2), pos=Em(5,30))

                self.bompos = wx.grid.Grid(self.panel2, -1, pos=Em(3,31), size=Em(27,5,1,0))
                self.bompos.DisableDragColSize()
                self.bompos.DisableDragRowSize()
                self.bompos.CreateGrid(5, 2)
                self.bompos.DisableDragGridSize()
                self.bompos.ShowScrollbars(wx.SHOW_SB_NEVER,wx.SHOW_SB_NEVER)
                self.bompos.SetColLabelValue(0, 'BOM/POS')
                self.bompos.SetColLabelValue(1, 'Filename')
                self.bompos.SetRowLabelSize(1)
                self.bompos.SetColSize(0, Em(9,1)[0])
                self.bompos.SetColSize(1, Em(18,1)[0])
                bomposfile = ['Bom-Top', 'Bom-Bottom', 'Pos-Top', 'Pos-Bottom']
                self.bompos.SetColLabelSize(Em(1,1)[1])
                for i in range(len(bomposfile)):
                    self.bompos.SetCellValue(i, 0, bomposfile[i])
                    self.bompos.SetReadOnly(i, 0, True)
                    self.bompos.SetRowSize(i, Em(1,1)[1])
                self.opt_BomMergeSide = wx.CheckBox(self.panel2, -1, 'BomMergeSide', pos=Em(31,32))
                self.opt_BomIncludeTHT = wx.CheckBox(self.panel2, -1, 'BomIncludeTHT', pos=Em(31,33))
                self.opt_PosMergeSide = wx.CheckBox(self.panel2, -1, 'PosMergeSide', pos=Em(31,34))
                self.opt_PosIncludeTHT = wx.CheckBox(self.panel2, -1, 'PosIncludeTHT', pos=Em(31,35))
                wx.StaticText(self.panel2, -1, getstr('DESC2'), pos=Em(2,37))
                self.Bind(wx.EVT_SIZE, self.OnSize)
                self.Bind(wx.EVT_CHECKBOX, self.OnCheck)
                rct = self.pluginSettings.get("Recent", False)
                if rct in self.pluginSettings.get("Recent", []):
                    self.manufacturers.SetStringSelection(rct)
                    self.Select(rct)
                self.opt_RefillExec.SetValue(self.pluginSettings.get('RefillExec',True))
                self.opt_DrcExec.SetValue(self.pluginSettings.get('DrcExec', True))
                self.opt_FabExec.SetValue(self.pluginSettings.get('FabExec', True))
                self.opt_BomPosExec.SetValue(self.pluginSettings.get('BomPosExec', True))

            def Ignore(self, e):
                wx.PostEvent(e.GetEventObject().GetParent().GetEventHandler(), e)

            def ClosePlugin(self):
                print(f'ClosePlugin : {execDialog}')
                json.dump(self.pluginSettings, open(self.settings_fname, "w"), indent=4)
#                if execDialog != None:
#                    execDialog.Close(None)

            def OnSize(self, e):
#                print("OnSize")
                self.clSize = self.GetClientSize()
                self.panel.SetSize(self.clSize.x, self.clSize.y)
                self.panel2.SetSize(self.clSize.x, self.clSize.y - Em(0,10)[1])

            def OnCheck(self, e):
#                print('OnCheck')
                self.pluginSettings['RefillExec'] = self.opt_RefillExec.GetValue()
                self.pluginSettings['DrcExec'] = self.opt_DrcExec.GetValue()
                self.pluginSettings['FabExec'] = self.opt_FabExec.GetValue()
                self.pluginSettings['BomPosExec'] = self.opt_BomPosExec.GetValue()

            def Set(self, settings):
                self.settings=dict(default_settings,**settings)
                l = self.settings.get('Layers',{})
                for i in range(self.layer.GetNumberRows()):
                    k = self.layer.GetCellValue(i, 0)
                    if l.get(k,None) != None:
                        self.layer.SetCellValue(i, 1, l.get(k))
                    else:
                        self.layer.SetCellValue(i, 1, '')
                l = self.settings.get('Drill',{})
                for i in range(self.drill.GetNumberRows()):
                    k = self.drill.GetCellValue(i,0)
                    if l.get(k,None) != None:
                        self.drill.SetCellValue(i, 1, l.get(k))
                    else:
                        self.drill.SetCellValue(i, 1, '')

                self.bompos.SetCellValue(0, 1, self.settings.get('BomFile',{}).get('TopFilename'))
                self.bompos.SetCellValue(1, 1, self.settings.get('BomFile',{}).get('BottomFilename'))
                self.bompos.SetCellValue(2, 1, self.settings.get('PosFile',{}).get('TopFilename'))
                self.bompos.SetCellValue(3, 1, self.settings.get('PosFile',{}).get('BottomFilename'))

                self.opt_PlotBorderAndTitle.SetValue(self.settings.get('PlotBorderAndTitle',False))
                self.opt_PlotFootprintValues.SetValue(self.settings.get('PlotFootprintValues',True))
                self.opt_PlotFootprintReferences.SetValue(self.settings.get('PlotFootprintReferences',True))
                self.opt_UseAuxOrigin.SetValue(self.settings.get('UseAuxOrigin', False))
                self.opt_SubtractMaskFromSilk.SetValue(self.settings.get('SubtractMaskFromSilk', False))
                self.opt_UseExtendedX2format.SetValue(self.settings.get('UseExtendedX2format', False))
                self.opt_CoordinateFormat46.SetValue(self.settings.get('CoordinateFormat46',True))
                self.opt_IncludeNetlistInfo.SetValue(self.settings.get('IncludeNetlistInfo',False))
                self.opt_DrillUnit.SetSelection(1 if self.settings.get('DrillUnitMM',True) else 0)
                self.opt_MirrorYAxis.SetValue(self.settings.get('MirrorYAxis', False))
                self.opt_MinimalHeader.SetValue(self.settings.get('MinimalHeader', False))
                self.opt_MergePTHandNPTH.SetValue(self.settings.get('MergePTHandNPTH', False))
                self.opt_RouteModeForOvalHoles.SetValue(self.settings.get('RouteModeForOvalHoles', True))
                zeros = self.settings.get('ZerosFormat',{})
                i = 0
                for k in zeros:
                    if(zeros[k]):
                        i = {'DecimalFormat':0,'SuppressLeading':1,'SuppressTrailing':2,'KeepZeros':3}.get(k,0)
                self.opt_ZerosFormat.SetSelection(i)

                map = self.settings.get('MapFileFormat',{})
                i = 4
                for k in map:
                    if(map[k]):
                        i = {'PostScript':0,'Gerber':1,'DXF':2,'SVG':3,'PDF':4}.get(k,4)
                self.opt_MapFileFormat.SetSelection(i)
                files=self.settings.get('OptionalFiles',[])
                if len(files)==0:
                    files=[{'name':'','content':''}]
                self.opt_OptionalFile.SetValue(files[0]['name'])
                self.opt_OptionalContent.SetValue(files[0]['content'])
                bom = self.settings.get('BomFile',{})
                pos = self.settings.get('PosFile',{})
                self.opt_BomMergeSide.SetValue(1 if bom.get('MergeSide',False) else 0)
                self.opt_BomIncludeTHT.SetValue(1 if bom.get('IncludeTHT',False) else 0)
                self.opt_PosMergeSide.SetValue(1 if pos.get('MergeSide',False) else 0)
                self.opt_PosIncludeTHT.SetValue(1 if pos.get('IncludeTHT',False) else 0)

            def Get(self):
                l = self.settings.get('Layers',{})
                for i in range(self.layer.GetNumberRows()):
                    k = self.layer.GetCellValue(i, 0)
                    v = self.layer.GetCellValue(i, 1)
                    l[k] = v
                self.settings['Layers'] = l
                d = self.settings.get('Drill',{})
                for i in range(self.drill.GetNumberRows()):
                    k = self.drill.GetCellValue(i, 0)
                    v = self.drill.GetCellValue(i, 1)
                    d[k] = v
                self.settings['Drill'] = d
                self.settings['PlotBorderAndTitle'] = self.opt_PlotBorderAndTitle.GetValue()
                self.settings['PlotFootprintValues'] = self.opt_PlotFootprintValues.GetValue()
                self.settings['PlotFootprintReferences'] = self.opt_PlotFootprintReferences.GetValue()
                self.settings['UseAuxOrigin'] = self.opt_UseAuxOrigin.GetValue()
                self.settings['SubtractMaskFromSilk'] = self.opt_SubtractMaskFromSilk.GetValue()
                self.settings['UseExtendedX2format'] = self.opt_UseExtendedX2format.GetValue()
                self.settings['CoordinateFormat46'] = self.opt_CoordinateFormat46.GetValue()
                self.settings['IncludeNetlistInfo'] = self.opt_IncludeNetlistInfo.GetValue()
                self.settings['DrillUnitMM'] = True if self.opt_DrillUnit.GetSelection() else False
                self.settings['MirrorYAxis'] = self.opt_MirrorYAxis.GetValue()
                self.settings['MinimalHeader'] = self.opt_MinimalHeader.GetValue()
                self.settings['MergePTHandNPTH'] = self.opt_MergePTHandNPTH.GetValue()
                self.settings['RouteModeForOvalHoles'] = self.opt_RouteModeForOvalHoles.GetValue()
                zeros = self.settings['ZerosFormat']
                i = self.opt_ZerosFormat.GetSelection()
                zeros['DecimalFormat'] = i == 0
                zeros['SuppressLeading'] = i == 1
                zeros['SuppressTrailing'] = i == 2
                zeros['KeepZeros'] = i == 3
                map = self.settings['MapFileFormat']
                i = self.opt_MapFileFormat.GetSelection()
                map['PostScript'] = i == 0
                map['Gerber'] = i == 1
                map['DXF'] = i == 2
                map['SVG'] = i == 3
                map['PDF'] = i == 4
                f = {'name':self.opt_OptionalFile.GetValue(), 'content':self.opt_OptionalContent.GetValue()}
                self.settings['OptionalFiles'] = [f]
                bom = self.settings['BomFile']
                pos = self.settings['PosFile']
                bom['TopFilename'] = self.bompos.GetCellValue(0, 1)
                bom['BottomFilename'] = self.bompos.GetCellValue(1, 1)
                pos['TopFilename'] = self.bompos.GetCellValue(2, 1)
                pos['BottomFilename'] = self.bompos.GetCellValue(3, 1)
                bom['MergeSide'] = self.opt_BomMergeSide.GetValue()
                bom['IncludeTHT'] = self.opt_BomIncludeTHT.GetValue()
                pos['MergeSide'] = self.opt_PosMergeSide.GetValue()
                pos['IncludeTHT'] = self.opt_PosIncludeTHT.GetValue()

#                bom['THT'] = self.opt_BOMTHT.GetValue()
#                bom['SMD'] = self.opt_BOMSMD.GetValue()
#                pos['THT'] = self.opt_PosTHT.GetValue()
#                pos['SMD'] = self.opt_PosSMD.GetValue()
                return self.settings

            def Select(self, name):
                print('Select:"'+name+'"')
                self.settings = default_settings
                if(name in self.json_data):
                    self.settings = self.json_data[name]
                else:
                    for k in self.json_data.keys():
                        self.settings = self.json_data[k]
                        break
                prefix_path = os.path.join(os.path.dirname(__file__))
                self.pluginSettings['Recent'] = name
                self.label.SetLabel(self.settings.get('Description', ''))
                self.url.SetValue(self.settings.get('URL','---'))
                self.gerberdir.SetValue(self.settings.get('GerberDir','Gerber'))
                self.zipfilename.SetValue(self.settings.get('ZipFilename','*.ZIP'))
                self.Set(self.settings)

            def PrepareDirs(self, mktemp):
                try:
                    doc = kicad.get_open_documents(DocumentType.DOCTYPE_PCB)
                except Exception as err:
                    alert(f'eror \n\n {err}')
                self.work_dir = doc[0].project.path
#                self.work_dir = os.path.dirname(__file__)
                self.temp_dir = os.path.join(self.work_dir, 'temp')
                self.board_path = os.path.join(self.temp_dir, '$$$.kicad_pcb')
                self.basename = os.path.splitext(os.path.basename(board.name))[0]
                self.gerber_dir = os.path.join(self.work_dir, self.gerberdir.GetValue())
                if not os.path.exists(self.gerber_dir):
                    os.makedirs(self.gerber_dir)
                if mktemp and not os.path.exists(self.temp_dir):
                    os.makedirs(self.temp_dir)

            def OnManufacturers(self,e):
                obj = e.GetEventObject()
                self.Select(obj.GetStringSelection())
                e.Skip()

            def OnDetail(self, e):
                if self.detailbtn.GetValue():
                    self.SetClientSize(wx.Size(self.szPanel[1][0], self.szPanel[1][1]))
                else:
                    self.SetClientSize(wx.Size(self.szPanel[0][0], self.szPanel[0][1]))
                if e:
                    e.Skip()

            def OnClose(self, e):
                e.Skip()
                self.Close()
        
            def OnExec(self, e):
                global report, execDialog
                report = 'Exec Start\n'
                try:
                    execDialog = ExecProgress(self, 'test')
                    if execDialog == None:
                        return
                    print('Exec : '+board.name)
                    execDialog.Progress(report)
                    self.PrepareDirs(True)
                    board.save_as(self.board_path, overwrite = True, include_project = False)
                    execDialog.Progress(report)
                    report += '\n--- Refill ---\n'
                    execDialog.Progress(report)
                    if self.pluginSettings['RefillExec']:
                        RefillExec(self, board)
                    else:
                        report += '    Skip\n'
                    report += '\n--- DRC ---\n'
                    execDialog.Progress(report)
                    if(self.pluginSettings['DrcExec']):
                        DrcExec(self, board)
                    else:
                        report += '    Skip\n'
                    report += '\n--- Gerber ---\n'
                    execDialog.Progress(report)
                    GerberExec(self, board)
                    report += '\n--- Fab ---\n'
                    execDialog.Progress(report)
                    if(self.pluginSettings['FabExec']):
                        FabExec(self, board)
                    else:
                        report += '    Skip\n'
                    report += '\n--- BOM/POS ---\n'
                    execDialog.Progress(report)
                    if(self.pluginSettings['BomPosExec']):
                        BomPosExec(self, board)
                    else:
                        report += '    Skip\n'
                    print('\nComplete\n\n')
                    report += '    \nComplete\n\n'
                    execDialog.Progress(report)
                    shutil.rmtree(self.temp_dir)
                    execDialog.Progress(report)
                except Exception as e:
                    execDialog.Complete(1)
                    print(e)
                    return
                execDialog.Complete(0)
#                execDialog.Close()
#                alert(report)
                e.Skip()
        try:
            board = GetBoard()
        except BaseException as e:
            exit()
        dialog = Dialog(None)
        dialog.OnDetail(None)
        dialog.Center()
        dialog.ShowModal()
        dialog.Destroy()


if __name__=='__main__':
    app = wx.App()
    gerberzipper = GerberZipper2()
    gerberzipper.Run()
