import os,zipfile

def addfile(zf, file):
    print(file)
    zf.write(file, file)

with zipfile.ZipFile('pcmpackage.zip', 'w', compression=zipfile.ZIP_DEFLATED) as zf:
    addfile(zf, 'metadata.json')
    addfile(zf, 'resources/icon.png')
    addfile(zf, 'plugins/gerber_zipper_2_action.py')
    addfile(zf, 'plugins/plugin.json')
    addfile(zf, 'plugins/requirements.txt')
    files = os.listdir('plugins/Locale')
    for file in files:
        if file.find('Test') != 0:
            addfile(zf, 'plugins/Locale/' + file)
    files = os.listdir('plugins/Manufacturers')
    for file in files:
        if file.find('Test') != 0:
            addfile(zf, 'plugins/Manufacturers/' + file)
    files = os.listdir('plugins/Assets/')
    for file in files:
        if file.find('Test') != 0:
            addfile(zf, 'plugins/Assets/' + file)
    print('pcmpackage.zip complete')
