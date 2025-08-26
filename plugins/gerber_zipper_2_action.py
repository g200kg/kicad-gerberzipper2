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


title = "GerberZipper2(Beta)"
version = "2.0.1"
chsize = (10,20)
strtab = {}
mainDialog = None
execDialog = None
board = None
startupinfo = None
kicadcli_path = ''
temp_basename = '__PCB__'

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
  "DrcSchematicPatity":False,
  "DrcAllTrackErrors":False,
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
  "FabFile":{
    "TopFilename":"*-Fab-Top.pdf",
    "BottomFilename":"",
    "TopLayers":"Edge.Cuts,F.Fab,Cmts.User,Dwgs.User",
    "BottomLayers":"Edge.Cuts,B.Fab,Cmts.User,Dwgs.User"
  },
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
    def __init__(self, parent):
#        alert(f'{self}, {parent}')
        wx.Dialog.__init__(self, parent, -1, f'{title} Exec', style=wx.CAPTION|wx.STAY_ON_TOP)
        self.SetClientSize(Em(60,20))
        panel = wx.Panel(self)
        self.text = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, pos=Em(1,1),size=Em(58,16))
        self.button = wx.Button(panel, wx.ID_OK, "Close", pos=Em(5,18),size=Em(10,1.5))
        self.button.Hide()
        self.button.Bind(wx.EVT_BUTTON,self.OnClose)
        self.Bind(wx.EVT_CLOSE,self.Close)
        self.board = board
        self.parent = parent
        self.Show()
        self.step = 0
        self.report = 'Start\n'
        self.Progress()
    def Progress(self):
        self.text.SetValue(self.report)
        self.text.ShowPosition(self.text.GetLastPosition())
        self.Refresh()
        self.Update()
        self.Raise()
        self.SetFocus()
        self.step += 1
        wx.CallLater(300, self.Exec)
    def OnClose(self, e):
        global execDialog
        print('ExecDialog.Close')
        self.EndModal(0)
        self.Destroy()
        execDialog = None
    def Complete(self, stat):
        self.button.Show()
        if not stat:
            wx.Bell()
    def Exec(self):
        try:
            if self.step == 1:
                self.report += '\n--- Prepare ---\n'
                self.Progress()
            elif self.step == 2:
                self.parent.PrepareDirs()
                self.board.save_as(self.parent.board_path, overwrite = True, include_project = False)
                self.report += '\nDone\n\n--- Refill ---\n'
                self.Progress()
            elif self.step == 3:
                if self.parent.pluginSettings['RefillExec']:
                    ret = RefillExec(self.parent, self.board)
                    print(ret)
                    self.report += ret['str']
                    if ret['stat'] > 0:
                        raise
                else:
                    self.report += 'Skip\n'
                self.report += '\n--- DRC ---\n'
                self.Progress()
            elif self.step == 4:
                if self.parent.pluginSettings['DrcExec']:
                    ret = DrcExec(self.parent, self.board)
#                    alert(f'Drc:{ret}')
                    self.report += ret['str']
                    if ret['stat'] > 0:
                        raise Exception(f'DRC Error : {ret["str"]}')
                else:
                    self.report += 'Skip\n'
                self.report += '\n--- Gerber ---\n'
                self.Progress()
            elif self.step == 5:
                ret = GerberExec(self.parent, self.board)
                self.report += ret['str']
                if ret['stat'] > 0:
                    raise
                self.report += '\n--- Fab ---\n'
                self.Progress()
            elif self.step ==6:
                if(self.parent.pluginSettings['FabExec']):
                    ret = FabExec(self.parent, self.board)
                    print(ret)
                    self.report += ret['str']
                    if ret['stat'] > 0:
                        raise
                else:
                    self.report += 'Skip\n'
                self.report += '\n--- Bom/Pos ---\n'
                self.Progress()
            elif self.step == 7:
                if(self.parent.pluginSettings['BomPosExec']):
                    ret = BomPosExec(self.parent, self.board)
                    self.report += ret['str']
                    if ret['stat'] > 0:
                        raise
                else:
                    self.report += 'Skip\n'
                self.Progress()
            elif self.step == 8:
                self.report += '\n\nComplete\n\n'
#                shutil.rmtree(self.parent.temp_dir)
                self.Complete(0)
                self.Progress()
        except Exception as e:
            alert(str(e), wx.ICON_ERROR)
            self.Complete(1)
#    self.Exec()
#    self.ShowModal()
#    self.Destroy()

#def ExecProgress(self, s, icon=0):
#    if self.FindWindowByName(f'{title} Exec') is None:
#        return ExecDialog()
#    else:
#        return None

def alert(s, icon=0):
    dialog = wx.MessageDialog(None, s, f'{title}', icon)
    r = dialog.ShowModal()
    return r

def InitEm():
    global chsize
    dc = wx.ScreenDC()
    font = wx.Font(pointSize=9,family=wx.DEFAULT,style=wx.NORMAL,weight=wx.NORMAL)
    dc.SetFont(font)
    tx = dc.GetTextExtent("M")
    chsize = (tx[0],tx[1]*1.6)
    if sys.platform == 'darwin':
        chsize = (tx[0],tx[1]*3)
#        alert(f'chsize:{chsize}\nsys.platform:{sys.platform}')

def SText(p, id, str, pos, size, style=wx.TE_LEFT):
    return wx.StaticText(p, id, str, pos=(pos[0],pos[1]+Em(0,0.15)[1]), size=size, style=style)

def Em(x,y):
    return (int(chsize[0]*x), int(chsize[1]*y))

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
    m = re.findall(r'\${([0-9a-zA-Z|_]*)}',s)
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
                alert(f'File creation error \n\n File : {fn}\n', wx.ICON_ERROR)
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
                    alert(f'File creation eror \n\n File : {fn}\n', wx.ICON_ERROR)
                    raise
            else:
                try:
                    self.f = open(self.fname, mode='w', encoding='utf-8')
                except Exception as err:
                    alert(f'File creation eror \n\n File : {fn}\n', wx.ICON_ERROR)
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

def SubprocRun(cmd):
    if startupinfo == None:
        ret = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        return ret.stdout
    else:
        ret = subprocess.run(cmd, capture_output=True, text=True, shell=True, startupinfo=startupinfo)
#        alert(f'ret.stdout:{ret.stdout}')
#        alert(f'ret.stderr:{ret.stderr}')
#        alert(f'retcode:{ret.returncode}')
        return ret.stdout

def GetBoard():
    global kicad
    global board
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
#    print('test')
#    print(DocumentType.items())
#    print('\n')
#    d = kicad.get_open_documents(DocumentType.DOCTYPE_UNKNOWN)
#    d = kicad.get_open_documents(DocumentType.DOCTYPE_PCB)
#    d = kicad.get_open_documents(DocumentType.DOCTYPE_SCHEMATIC)
#    d = kicad.get_open_documents(DocumentType.DOCTYPE_PROJECT)
#    print(d)
#    print(DocumentType.items())
#    print('get_open_documents ok')
#    print(f'type:{type(d)}')
#    print(f'd:{d}')
#    exit()
    try:
        board = kicad.get_board()
    except BaseException as e:
        alert(f"PCB Open Error : ({e})", wx.ICON_ERROR)
        exit()
    return board

def GerberExec(self, board):
    self.settings = self.Get()
    zipfiles = []

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
    cmd = f'"{kicadcli_path}" pcb export gerbers {opt} --no-protel-ext -o {self.temp_dir} -l {lystr} {self.board_path}'
    print(f'cmd : {cmd}')
    ret = SubprocRun(cmd)
    for k in kylist:
        if layers[k] != '':
            fname = getfname(k)
            renamefile(self.temp_dir, f'{temp_basename}-{fname}.gbr', self.gerber_dir, f'{layers[k].replace("*",self.basename)}', zipfiles)

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
    cmd = f'"{kicadcli_path}" pcb export drill {opt} -o {self.temp_dir} --format excellon {self.board_path}'
    print(f'cmd : {cmd}')
    ret = SubprocRun(cmd)
    print(f'4:{ret}')
    if self.settings['MergePTHandNPTH']:
        renamefile(self.temp_dir, f'{temp_basename}.drl', self.gerber_dir, self.settings['Drill']['Drill'].replace('*', self.basename), zipfiles)
        if self.settings['Drill']['DrillMap'] != '':
            renamefile(self.temp_dir, f'{temp_basename}-drl_map.{mext}', self.gerber_dir, self.settings['Drill']['DrillMap'].replace('*', self.basename), zipfiles)
    else:
        renamefile(self.temp_dir, f'{temp_basename}-PTH.drl', self.gerber_dir, self.settings['Drill']['Drill'].replace('*', self.basename), zipfiles)
        renamefile(self.temp_dir, f'{temp_basename}-NPTH.drl', self.gerber_dir, self.settings['Drill']['NPTH'].replace('*', self.basename), zipfiles)
        if self.settings['Drill']['DrillMap'] != '':
            renamefile(self.temp_dir, f'{temp_basename}-PTH-drl_map.{mext}', self.gerber_dir, self.settings['Drill']['DrillMap'].replace('*', self.basename), zipfiles)
        if self.settings['Drill']['NPTHMap'] != '':
            renamefile(self.temp_dir, f'{temp_basename}-NPTH-drl_map.{mext}', self.gerber_dir, self.settings['Drill']['NPTHMap'].replace('*', self.basename), zipfiles)

    # Zip
    zipfname = f'{self.gerber_dir}/{self.zipfilename.GetValue().replace("*",self.basename)}'
    with zipfile.ZipFile(zipfname,'w',compression=zipfile.ZIP_DEFLATED) as f:
        for i in range(len(zipfiles)):
            fnam = zipfiles[i]
            if os.path.exists(fnam):
                f.write(fnam, os.path.basename(fnam))
    print('zip ok')
    return {'stat':0, 'str':f'Output : {zipfname}\nDone\n'}

def RefillExec(self, board):
    print('RefillExec')
    print(board)
    r = board.refill_zones(block=True)
    print(r)
    return {'stat':0, 'str':'Done\n'}

def DrcExec(self, board):
    print('DrcExec')
    ret = ''
    opt = ''
    if self.settings.get('DrcSchematicParity',False):
        opt += '--schematic-parity '
    if self.settings.get('DrcAllTrackErrors',False):
        opt += '--all-track-errors '
    cmd = f'"{kicadcli_path}" pcb drc {opt} -o {self.gerber_dir}/{self.basename}.rpt --format report {self.board_path}'
#    cmd = f'kicad-cli pcb drc {opt} -o {self.gerber_dir}/{self.basename}.rpt --format json {self.board_path}'
#    alert(f'cmd:{cmd}')
    print(f'cmd : {cmd}')
    ret = SubprocRun(cmd)
#    alert(f'ret:{ret}')
    nums = re.findall(r' (\d*) ', ret)
#    print(f'nums:{nums}')
    err = 0
    for n in nums:
        if int(n) > 0:
            err = 1
    r = {'stat':err, 'str':ret}
    print(f'r:{r}')
    return r

def FabExec(self, board):
    print('FabExec')
    ret = ''
    layers = ''
    fabfile = self.settings['FabFile']
    fn = fabfile['TopFilename']
    if fn != '':
        layers = fabfile['TopLayers']
        cmd = f'"{kicadcli_path}" pcb export pdf -o {self.gerber_dir}/{fn.replace("*",self.basename)} -l {layers} {self.board_path}'
        print(f'cmd : {cmd}')
        res = SubprocRun(cmd)
#        print(f'returncode : {res.returncode}')
#        print(f'stderr : {res.stderr}')
        ret += res
    fn = fabfile['BottomFilename']
    if fn != '':
        layers = fabfile['BottomLayers']
        cmd = f'"{kicadcli_path}" pcb export pdf -o {self.gerber_dir}/{fn.replace("*",self.basename)} -l {layers} {self.board_path}'
        print(f'cmd : {cmd}')
        ret += SubprocRun(cmd)
    return {'stat':0, 'str':ret}

def BomPosExec(self, board):
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
        return {'stat':1, 'str':''}
    return {'stat':0, 'str':f'BomTop:{bom_fnameT}\nBomBottom:{bom_fnameB}\nPosTop:{pos_fnameT}\nPosBottom:{pos_fnameB}\nDone'}

class GerberZipper2():
    def Run(self):
        global mainDialog,board
        class Dialog(wx.Dialog):
            def __init__(self, parent):
                global strtab
                global startupinfo
                global kicadcli_path

                kicadcli_path = kicad.get_kicad_binary_path('kicad-cli')
                prefix_path = os.path.join(os.path.dirname(__file__))
                self.pluginSettings = {}
                self.icon_file_name = os.path.join(os.path.dirname(__file__), 'Assets/icon48.png')
                try:
                    doc = kicad.get_open_documents(DocumentType.DOCTYPE_PCB)
                except Exception as err:
                    alert(f'eror \n\n {err}', wx.ICON_ERROR)
                self.work_dir = doc[0].project.path
                self.temp_dir = os.path.join(self.work_dir, '~gerberzipper_temp')
                print(f'temp_dir:{self.temp_dir}')
                if os.path.exists(self.temp_dir):
                    if alert('It may already be running. Run anyway?', wx.OK|wx.CANCEL) == wx.ID_CANCEL:
                        exit(0)
                else:
                    os.makedirs(self.temp_dir)
                atexit.register(self.ClosePlugin)
                self.manufacturers_dir = os.path.join(os.path.dirname(__file__), 'Manufacturers')
                manufacturers_list = glob.glob(f'{self.manufacturers_dir}/*.json')
                self.json_data = {}
                for fname in manufacturers_list:
                    try:
                        d = json.load(codecs.open(fname, 'r', 'utf-8'))
                        self.json_data[d['Name']] = json.load(codecs.open(fname, 'r', 'utf-8'))
                    except Exception as err:
                        tb = sys.exc_info()[2]
                        alert(f'JSON error \n\n File : {os.path.basename(fname)}\n{err.with_traceback(tb)}', wx.ICON_WARNING)
                self.settings_fname = os.path.join(prefix_path, 'gerberzipper2.json')
                if os.path.exists(self.settings_fname):
                    self.pluginSettings = json.load(open(self.settings_fname))
                else:
                    self.pluginSettings["Recent"] = next(iter(self.json_data))
                    self.pluginSettings["RefillExec"] = True
                    self.pluginSettings["DrcExec"] = False
                    self.pluginSettings["FabExec"] = True
                    self.pluginSettings["BomPosExec"] = True
                self.locale_dir = os.path.join(os.path.dirname(__file__), "Locale")
                locale_list = glob.glob(f'{self.locale_dir}/*.json')
                strtab = {}
                for fpath in locale_list:
                    fname = os.path.splitext(os.path.basename(fpath))[0]
                    strtab[fname] = json.load(codecs.open(fpath, 'r', 'utf-8'))
                if hasattr(subprocess, 'STARTUPINFO'):
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags = subprocess.CREATE_NEW_CONSOLE | subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                InitEm()
                self.szPanel = [Em(75,9.5), Em(75,25)]
                wx.Dialog.__init__(self, parent, id=-1, title=f'{title} {version}', style=wx.DEFAULT_DIALOG_STYLE | wx.CENTRE | wx.RESIZE_BORDER)
                self.panel = wx.Panel(self)
                self.SetIcon(wx.Icon(self.icon_file_name))

                manufacturers_arr=[]
                for item in self.json_data.keys():
                    manufacturers_arr.append(item)
                SText(self.panel, -1, 'Make Gerbers/BOM/POS/ZIP for specific PCB manufacturers. (IPC API)', pos=Em(1,0.5), size=Em(70,1))
                SText(self.panel, -1, 'Manufacturers', pos=Em(1,1.5), size=Em(14,1), style=wx.TE_RIGHT)
                self.manufacturers = wx.ComboBox(self.panel, -1, 'Select Manufacturers', pos=Em(16,1.5), size=Em(30,1.5), choices=manufacturers_arr, style=wx.CB_READONLY)

                SText(self.panel, -1, 'URL', pos=Em(1,3), size=Em(14,1), style=wx.TE_RIGHT)
                self.url = wx.TextCtrl(self.panel, -1, '', pos=Em(16,3), size=Em(30,1), style=wx.TE_READONLY)

                SText(self.panel, -1, 'Gerber Dir', pos=Em(1,4), size=Em(14,1), style=wx.TE_RIGHT)
                self.gerberdir = wx.TextCtrl(self.panel, -1, '', pos=Em(16,4), size=Em(30,1))

                SText(self.panel, -1, 'Zip Filename', pos=Em(1,5), size=Em(14,1), style=wx.TE_RIGHT)
                self.zipfilename = wx.TextCtrl(self.panel, -1, '', pos=Em(16,5), size=Em(30,1))
                SText(self.panel, -1, 'Description', pos=Em(1,6), size=Em(14,1), style=wx.TE_RIGHT)
                self.label = SText(self.panel, -1, '',pos=Em(16,6), size=Em(45,1))

                self.opt_RefillExec = wx.CheckBox(self.panel, -1, 'Refill Zones', pos=Em(52,2), size=Em(15,1))
                self.opt_DrcExec = wx.CheckBox(self.panel, -1, 'Exec DRC', pos=Em(52,3), size=Em(15,1))
                self.opt_FabExec = wx.CheckBox(self.panel, -1, 'Generate Fab', pos=Em(52,4), size=Em(15,1))
                self.opt_BomPosExec = wx.CheckBox(self.panel, -1, 'Generate BOM/POS', pos=Em(52,5), size=Em(15,1))

                self.detailbtn = wx.ToggleButton(self.panel, -1, 'Show Settings Details', pos=Em(2,7.5), size=Em(18,1.5))
                self.execbtn = wx.Button(self.panel, -1, 'Make Gerber and ZIP', pos=Em(21,7.5), size=Em(18,1.5))
                self.clsbtn = wx.Button(self.panel, -1, 'Close', pos=Em(40,7.5), size=Em(18,1.5))
                self.manufacturers.Bind(wx.EVT_COMBOBOX, self.OnManufacturers)
                self.detailbtn.Bind(wx.EVT_TOGGLEBUTTON, self.OnDetail)
                self.execbtn.Bind(wx.EVT_BUTTON, self.OnExec)
                self.clsbtn.Bind(wx.EVT_BUTTON, self.OnClose)

                self.panel2 = wx.lib.scrolledpanel.ScrolledPanel(self.panel, -1, pos=Em(0,10), size=Em(75,10.5), style=wx.BORDER_SUNKEN)
                self.panel2.SetScrollbars(20,20,0,100)

                stxt = SText(self.panel2, -1, 'DRC', style=wx.TE_CENTER, pos=Em(1,0.5), size=Em(12,1))
                stxt.SetBackgroundColour('#c0c0d0')

                self.opt_DrcSchematicParity = wx.CheckBox(self.panel2, -1, 'SchematicParity', pos=Em(24,2))
                self.opt_DrcAllTrackErrors = wx.CheckBox(self.panel2, -1, "AllTrackErrors", pos=Em(24,3))

                stxt = SText(self.panel2, -1, 'ZIP contents', style=wx.TE_CENTER, pos=Em(1,3.5), size=Em(12,1))
                stxt.SetBackgroundColour('#c0c0d0')
                wx.StaticBox(self.panel2, -1,'Gerber', pos=Em(2,5), size=Em(65,13))

                self.layer = wx.grid.Grid(self.panel2, -1, pos=Em(3,6), size=Em(19,11))
                self.layer.SetColLabelSize(Em(1,1)[1])
                self.layer.DisableDragColSize()
                self.layer.DisableDragRowSize()
                self.layer.CreateGrid(len(layer_list), 2)
                self.layer.SetColLabelValue(0, 'Layer')
                self.layer.SetColLabelValue(1, 'Filename')
                self.layer.SetRowLabelSize(1)
                self.layer.ShowScrollbars(wx.SHOW_SB_NEVER,wx.SHOW_SB_DEFAULT)
                self.layer.SetColSize(0, Em(9,1)[0])
                self.layer.SetColSize(1, Em(9,1)[0])
                for i in range(len(layer_list)):
                    self.layer.SetCellValue(i, 0, layer_list[i]['name'])
                    self.layer.SetReadOnly(i, 0, isReadOnly=True)
                self.opt_UseAuxOrigin = wx.CheckBox(self.panel2, -1, 'UseAuxOrigin', pos=Em(24,6.5))

                self.opt_PlotBorderAndTitle = wx.CheckBox(self.panel2, -1, 'PlotBorderAndTitle', pos=Em(24,9))
                self.opt_PlotFootprintValues = wx.CheckBox(self.panel2, -1, 'PlotFootprintValues', pos=Em(24,10))
                self.opt_PlotFootprintReferences = wx.CheckBox(self.panel2, -1, 'PlotFootprintReferences', pos=Em(24,11))
                self.opt_SubtractMaskFromSilk = wx.CheckBox(self.panel2, -1, 'SubtractMaskFromSilk', pos=Em(24, 12))
                self.opt_UseExtendedX2format = wx.CheckBox(self.panel2, -1, 'UseExtendedX2format', pos=Em(24, 13))
                self.opt_CoordinateFormat46 = wx.CheckBox(self.panel2, -1, 'CoordinateFormat46', pos=Em(24, 14))
                self.opt_IncludeNetlistInfo = wx.CheckBox(self.panel2, -1, 'IncludeNetlistInfo', pos=Em(24, 15))

                wx.StaticBox(self.panel2, -1,'Drill', pos=Em(2,18), size=Em(65,8))
                self.drill = wx.grid.Grid(self.panel2, -1, pos=Em(3,19), size=Em(19,6))
                self.drill.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)
                self.drill.DisableDragColSize()
                self.drill.DisableDragRowSize()
                self.drill.CreateGrid(5, 2)
                self.drill.DisableDragGridSize()
                self.drill.ShowScrollbars(wx.SHOW_SB_NEVER,wx.SHOW_SB_NEVER)
                self.drill.SetColLabelValue(0, 'Drill')
                self.drill.SetColLabelValue(1, 'Filename')
                self.drill.SetRowLabelSize(1)
                self.drill.SetColSize(0, Em(9,1)[0])
                self.drill.SetColSize(1, Em(10,1)[0])
                drillfile = ['Drill', 'DrillMap', 'NPTH', 'NPTHMap', 'Report']
                self.drill.SetColLabelSize(Em(1,1)[1])
                for i in range(len(drillfile)):
                    self.drill.SetCellValue(i, 0, drillfile[i])
                    self.drill.SetReadOnly(i, 0, True)
                    self.drill.SetRowSize(i, Em(1,1)[1])
                self.opt_MirrorYAxis = wx.CheckBox(self.panel2, -1, 'MirrorYAxis', pos=Em(24,20))
                self.opt_MinimalHeader = wx.CheckBox(self.panel2, -1, 'MinimalHeader', pos=Em(24,21))
                self.opt_MergePTHandNPTH = wx.CheckBox(self.panel2, -1, 'MergePTHandNPTH', pos=Em(24,22))
                self.opt_RouteModeForOvalHoles = wx.CheckBox(self.panel2, -1, 'RouteModeForOvalHoles', pos=Em(24,23))
                SText(self.panel2, -1, 'Drill Unit :', pos=Em(43,20), size=Em(10,1), style=wx.TE_RIGHT)
                self.opt_DrillUnit = wx.ComboBox(self.panel2, -1, '', choices=('inch','mm'), style=wx.CB_READONLY, pos=Em(54,20), size=Em(8,1))
                self.opt_DrillUnit.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)
                SText(self.panel2, -1, 'Zeros :', pos=Em(43,21.5), size=Em(10,1), style=wx.TE_RIGHT)
                self.opt_ZerosFormat = wx.ComboBox(self.panel2, -1, '', choices=('DecimalFormat','SuppressLeading','SuppresTrailing', 'KeepZeros'), pos=Em(54,21.5), size=Em(12,1), style=wx.CB_READONLY)
                self.opt_ZerosFormat.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)
                SText(self.panel2, -1, 'MapFileFormat :', pos=Em(43,23), size=Em(10,1), style=wx.TE_RIGHT)
                self.opt_MapFileFormat = wx.ComboBox(self.panel2, -1, '', choices=('PostScript','Gerber','DXF','SVG','PDF'), pos=Em(54,23), size=Em(8,1), style=wx.CB_READONLY)
                self.opt_MapFileFormat.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)

                wx.StaticBox(self.panel2, -1,'Other', pos=Em(2,26), size=Em(65,3))

                self.opt_OptionalLabel = SText(self.panel2, -1, 'OptionalFile:', pos=Em(4,27), size=Em(10,1))
                self.opt_OptionalFile = wx.TextCtrl(self.panel2, -1, '', pos=Em(15,27), size=Em(12,1))
                self.opt_OptionalContent = wx.TextCtrl(self.panel2, -1, '', pos=Em(28,27), size=Em(37,1))

                stxt = SText(self.panel2, -1, 'Fab PDF', style=wx.TE_CENTER, pos=Em(1, 29.5), size=Em(12,1))
                stxt.SetBackgroundColour('#c0c0d0')

                self.fab = wx.grid.Grid(self.panel2, -1, pos=Em(3,31), size=Em(55,3))
                self.fab.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)
                self.fab.DisableDragColSize()
                self.fab.DisableDragRowSize()
                self.fab.CreateGrid(2, 3)
                self.fab.DisableDragGridSize()
                self.fab.ShowScrollbars(wx.SHOW_SB_NEVER,wx.SHOW_SB_NEVER)
                self.fab.SetColLabelValue(0, 'Fab')
                self.fab.SetColLabelValue(1, 'Filename')
                self.fab.SetColLabelValue(2, 'Layers')
                self.fab.SetRowLabelSize(1)
                self.fab.SetColSize(0, Em(9,1)[0])
                self.fab.SetColSize(1, Em(18,1)[0])
                self.fab.SetColSize(2, Em(28,1)[0])
                fabfile = ['Fab-Top', 'Fab-Bottom']
                self.fab.SetColLabelSize(Em(1,1)[1])
                for i in range(len(fabfile)):
                    self.fab.SetCellValue(i, 0, fabfile[i])
                    self.fab.SetReadOnly(i, 0, True)
                    self.fab.SetRowSize(i, Em(1,1)[1])

                stxt = SText(self.panel2, -1, 'Bom/Pos', style=wx.TE_CENTER, pos=Em(1, 34.5), size=Em(12,1))
                stxt.SetBackgroundColour('#c0c0d0')

                self.bompos = wx.grid.Grid(self.panel2, -1, pos=Em(3,36), size=Em(27,5))
                self.bompos.Bind(wx.EVT_MOUSEWHEEL, self.Ignore)
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
                self.opt_BomMergeSide = wx.CheckBox(self.panel2, -1, 'BomMergeSide', pos=Em(31,37))
                self.opt_BomIncludeTHT = wx.CheckBox(self.panel2, -1, 'BomIncludeTHT', pos=Em(31,38))
                self.opt_PosMergeSide = wx.CheckBox(self.panel2, -1, 'PosMergeSide', pos=Em(31,39))
                self.opt_PosIncludeTHT = wx.CheckBox(self.panel2, -1, 'PosIncludeTHT', pos=Em(31,40))
                SText(self.panel2, -1, 'The changes here are temporary. If you want to change it permanently edit the corresponding json file', pos=Em(2,42), size=Em(80,1))
                self.Bind(wx.EVT_SIZE, self.OnSize)
                self.Bind(wx.EVT_CHECKBOX, self.OnCheck)
                rct = self.pluginSettings.get("Recent", None)
                if rct != None:
                    self.Select(rct)
                    self.manufacturers.SetSelection(self.manufacturers.FindString(rct))
                self.opt_RefillExec.SetValue(self.pluginSettings.get('RefillExec',True))
                self.opt_DrcExec.SetValue(self.pluginSettings.get('DrcExec', True))
                self.opt_FabExec.SetValue(self.pluginSettings.get('FabExec', True))
                self.opt_BomPosExec.SetValue(self.pluginSettings.get('BomPosExec', True))
                self.timer = wx.Timer(self)
                self.Bind(wx.EVT_TIMER, self.Watch)
                self.timer.Start(5000)

            def Watch(self, e):
                global mainDialog, execDialog
                try:
                    al = board.get_active_layer()
                    print(f'Watch ActiveLayer:{al}')
                except Exception as err:
                    if err.__class__.__name__ == 'ApiError':
                        print(f'Watch exception:{mainDialog}')
                        if execDialog != None:
                            execDialog.Destroy()
                            execDialog = None
                        if mainDialog != None:
                            mainDialog.Destroy()
                            mainDialog = None
                        exit(0)
                if not os.path.exists(self.temp_dir):
                    exit(0)


            def Ignore(self, e):
                wx.PostEvent(self.panel2.GetEventHandler(), e)

            def ClosePlugin(self):
                print('ClosePlugin')
                if hasattr(self, 'settings_fname'):
                    json.dump(self.pluginSettings, open(self.settings_fname, "w"), indent=4)
                if os.path.exists(self.temp_dir):
                    shutil.rmtree(self.temp_dir)

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

                self.opt_DrcSchematicParity.SetValue(self.settings.get('DrcSchematicParity',False))
                self.opt_DrcAllTrackErrors.SetValue(self.settings.get('DrcAllTrackErrors',False))

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

                self.fab.SetCellValue(0, 1, self.settings.get('FabFile',{}).get('TopFilename'))
                self.fab.SetCellValue(1, 1, self.settings.get('FabFile',{}).get('BottomFilename'))
                self.fab.SetCellValue(0, 2, self.settings.get('FabFile',{}).get('TopLayers'))
                self.fab.SetCellValue(1, 2, self.settings.get('FabFile',{}).get('BottomLayers'))

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
                self.settings['DrcSchematicParity'] = self.opt_DrcSchematicParity.GetValue()
                self.settings['DrcAllTrackErrors'] = self.opt_DrcAllTrackErrors.GetValue()
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

                fab = self.settings['FabFile']
                fab['TopFilename'] = self.fab.GetCellValue(0,1)
                fab['BottomFilename'] = self.fab.GetCellValue(1, 1)
                fab['TopLayers'] = self.fab.GetCellValue(0, 2)
                fab['BottomLayers'] = self.fab.GetCellValue(1, 2)

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
                return self.settings

            def Select(self, name):
                print('Select:"'+name+'"')
                self.settings = default_settings
                if(name in self.json_data):
                    self.settings = self.json_data[name]
#                else:
#                    for k in self.json_data.keys():
#                        self.settings = self.json_data[k]
#                        break
                prefix_path = os.path.join(os.path.dirname(__file__))
                self.pluginSettings['Recent'] = name
                self.label.SetLabel(self.settings.get('Description', ''))
                self.url.SetValue(self.settings.get('URL','---'))
                self.gerberdir.SetValue(self.settings.get('GerberDir','Gerber'))
                self.zipfilename.SetValue(self.settings.get('ZipFilename','*.ZIP'))
                self.Set(self.settings)

            def PrepareDirs(self):
#                try:
#                    doc = kicad.get_open_documents(DocumentType.DOCTYPE_PCB)
#                except Exception as err:
#                    alert(f'eror \n\n {err}', wx.ICON_ERROR)
#                self.work_dir = doc[0].project.path
#                self.temp_dir = os.path.join(self.work_dir, 'gerberzipper_temp')
                self.board_path = os.path.join(self.temp_dir, '__pcb__.kicad_pcb')
                self.basename = os.path.splitext(os.path.basename(board.name))[0]
                self.gerber_dir = os.path.join(self.work_dir, self.gerberdir.GetValue())
                if not os.path.exists(self.gerber_dir):
                    os.makedirs(self.gerber_dir)
#                if not os.path.exists(self.temp_dir):
#                    os.makedirs(self.temp_dir)

            def OnManufacturers(self,e):
                obj = e.GetEventObject()
                self.Select(obj.GetStringSelection())
                e.Skip()

            def OnDetail(self, e):
                if self.detailbtn.GetValue():
                    print(f'SetClientSize:{self.szPanel[1][0]}, {self.szPanel[1][1]}')
                    self.panel2.Scroll(0,0)
                    self.panel2.Refresh()
                    self.panel2.Update()
                    self.panel2.Show()
                    self.SetClientSize(wx.Size(self.szPanel[1][0], self.szPanel[1][1]))
                else:
                    print(f'SetClientSize:{self.szPanel[0][0]}, {self.szPanel[0][1]}')
                    self.panel2.Scroll(0,0)
                    self.panel2.Refresh()
                    self.panel2.Update()
                    self.panel2.Hide()
                    self.SetClientSize(wx.Size(self.szPanel[0][0], self.szPanel[0][1]))
                if e:
                    e.Skip()

            def OnClose(self, e):
                e.Skip()
                self.Close()
        
            def OnExec(self, e):
                global execDialog
                report = 'Exec Start\n'
                print(f'self:{self}')
                print(f'execD:{execDialog}')
                if execDialog == None:
                    execDialog = ExecDialog(mainDialog)
                    execDialog.ShowModal()

        board = GetBoard()
        mainDialog = Dialog(None)
        mainDialog.OnDetail(None)
        mainDialog.Center()
        mainDialog.ShowModal()
        if mainDialog != None:
            mainDialog.Destroy()



if __name__=='__main__':
    app = wx.App()
    gerberzipper = GerberZipper2()
    gerberzipper.Run()
