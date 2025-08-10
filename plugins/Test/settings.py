{
    'Name': 'Test',
    'Description': 'Make Gerber-files/Zip for Test',
    'URL': 'https: //www.pcbway.com',
    'GerberDir': 'Gerber',
    'ZipFilename': '*-Test.ZIP',
    'Layers': {
        'F.Cu': '*.gtl',
        'B.Cu': '*.gbl',
        'F.Paste': '*.gtp',
        'B.Paste': '*.gbp',
        'F.SilkS': '*.gto',
        'B.SilkS': '*.gbo',
        'F.Mask': '*.gts',
        'B.Mask': '*.gbs',
        'Edge.Cuts': '*.gm1',
        'In1.Cu': '*.gl2',
        'In2.Cu': '*.gl3',
        'In3.Cu': '*.gl4',
        'In4.Cu': '*.gl5',
        'In5.Cu': '*.gl6',
        'In6.Cu': '*.gl7',
        'Dwgs.User': '*.dwgs'
    },
    'PlotBorderAndTitle': False,
    'PlotFootprintValues': True,
    'PlotFootprintReferences': True,
    'ForcePlotInvisible': False,
    'ExcludeEdgeLayer': True,
    'ExcludePadsFromSilk': True,
    'DoNotTentVias': False,
    'UseAuxOrigin': False,
    'LineWidth': 0.1,
    'CoodinateFormat46': True,
    'SubtractMaskFromSilk': False,
    'UseExtendedX2format': False,
    'IncludeNetlistInfo': False,
    'Drill': {
        'Drill': '*.drl',
        'DrillMap': '',
        'NPTH': '',
        'NPTHMap': '',
        'Report': ''
    },
    'DrillUnitMM': True,
    'MirrorYAxis': False,
    'MinimalHeader': True,
    'MergePTHandNPTH': True,
    'RouteModeForOvalHoles': True,
    'ZerosFormat': {
        'DecimalFormat': False,
        'SuppressLeading': True,
        'SuppressTrailing': False,
        'KeepZeros': False
    },
    'MapFileFormat': {
        'PostScript': True,
        'Gerber': False,
        'DXF': False,
        'SVG': False,
        'PDF': False
    },
    'OptionalFiles': [],
    'BomFile': {
        'TopFilename': '*-PCBWay-BOM.csv',
        'BottomFilename': '',
        'MergeSide': True,
        'IncludeTHT': True,
        'Tabs': [8, 30, 8, 30, 30, 35, 35, 16],
        'Header': 'Item #,Designator,Qty,Manufacturer,Mfg Part #,Description/Value,Package/Footprint,Type',
        'Row': '${num},"${ref}",${qty},"${MFR}","${PN|val}","${val}","${fp}","${type}"'
    },
    'PosFile': {
        'TopFilename': '*-PCBWay-POS-Top.csv',
        'BottomFilename': '*-PCBWay-POS-Bottom.csv',
        'MergeSide': False,
        'IncludeTHT': False,
        'Tabs': [10, 10, 10, 6, 10, 36, 8],
        'Header': 'Designator,Footprint,Mid X,Mid Y,TB,Rotation,Comment',
        'Row': '"${ref}","${fp}",${x},${y},"${side1}",${rot},"${type}"'
    }
}
    