import os,zipfile

def addfile(zf, file):
    print(file)
    zf.write(file, file)

with zipfile.ZipFile('pcmpackage.zip', 'w', compression=zipfile.ZIP_DEFLATED) as zf:
    addfile(zf, 'metadata.json')
    addfile(zf, 'resources/icon32.png')
    addfile(zf, 'plugins/gerberzipper2.py')
    files = os.listdir('plugins/Locale')
    for file in files:
        addfile(zf, 'plugins/Locale/' + file)
    files = os.listdir('plugins/Manufacturers')
    for file in files:
        addfile(zf, 'plugins/Manufacturers/' + file)
    files = os.listdir('plugins/Assets/')
    for file in files:
        addfile(zf, 'plugins/Assets/' + file)
    print('pcmpackage.zip complete')
