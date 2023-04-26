from PyInstaller.utils.hooks import collect_submodules

import os.path
import wx

# Determine the location of our source files.
app_root = os.path.dirname(os.path.abspath(__file__))
mainwindow_path = os.path.join(app_root, 'MainWindow')
mendreport_path = os.path.join(app_root, 'MonthEndReport')

# Define the PyInstaller options.
opts = {
    'name': 'ExcelToolv2.0',
    'onefile': True,
    'hiddenimports': [
        'openpyxl',
        'wx',
    ],
    'additional_files': [
        (os.path.join(mendreport_path, 'MonthEndReport.py'), 'MonthEndReport'),
        (os.path.join(mendreport_path, 'SolarPower.py'), 'MonthEndReport'),
    ],
}

# Define the PyInstaller targets.
exe = [
    wx.Exe(
        script=os.path.join(mainwindow_path, 'MainWindow.py'),
        base=None,
        target_name='ExcelToolv2.0.exe',
        icon=None,
    ),
]

# Collect submodules for PyInstaller.
hiddenimports = collect_submodules('wx')

# Define the PyInstaller build details.
build_exe_options = {
    'packages': [],
    'excludes': [],
    'include_files': [],
    'hiddenimports': hiddenimports,
}

# Call PyInstaller to build the executable.
if __name__ == '__main__':
    from PyInstaller.building.build_main import main
    main(args=['--name', opts['name'], '-F', '--hidden-import', ','.join(opts['hiddenimports']), '--add-data', ','.join(opts['additional_files']), '--distpath', '.', '--workpath', 'build', '--noconfirm', os.path.join(mainwindow_path, 'MainWindow.py')], **build_exe_options)
