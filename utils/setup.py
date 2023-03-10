
from cx_Freeze import setup, Executable

# 打包的脚本文件路径
target_file = r"..\MainWindow\MainWindow.py"

# 打包后的exe文件保存路径
exe_path = 'ExcelTool.exe'

# 包含的模块和第三方库
includefiles = [
    "../MonthEndReport/MonthEndReport.py",
    "../MonthEndReport/SolarPower.py",
    "./manifest.xml"
]

includes = [
    "os",
    "time",
    "decimal",
    "wx",
    "openpyxl"
]

# 配置打包参数
options = {
    "build_exe": {
        "include_msvcr": True,
        "include_files": includefiles,
        "includes": includes,
    }
}

# 创建Executable对象
build_exe = Executable(
    script=target_file,
    base=None,
    targetName=exe_path
)

# 调用setup函数进行打包
setup(
    name='ExcelTool',
    version='1.0',
    description='An Excel tool',
    options=options,
    executables=[build_exe]
)
