from PyInstaller.utils.hooks import collect_all, collect_data_files

# 收集 pandas 模块的全部依赖和数据文件
datas, binaries, hiddenimports = collect_all('pandas')

# 补充常用子模块（保证 Excel 读写、日期、缺失值转换功能正常）
hiddenimports += [
    'pandas.core',
    'pandas.core.frame',
    'pandas.core.tools',
    'pandas.io',
    'pandas.io.excel',
    'pandas.io.excel._openpyxl',
    'pandas._libs.tslibs',
    'pandas._libs.missing',
]

datas += collect_data_files('pandas')

