from PyInstaller.utils.hooks import collect_all, collect_data_files

# 收集 pandas 所有模块和数据文件（基础操作）
datas, binaries, hiddenimports = collect_all('pandas')

# 补充你用到的功能依赖的关键子模块
# 1. pd.isna 依赖 core 模块
# 2. pd.ExcelFile、pd.read_excel 依赖 io.excel 模块（尤其是 openpyxl 引擎）
# 3. pd.DataFrame 依赖 core.frame 模块
# 4. pd.to_numeric、pd.to_datetime 依赖 core.tools 和 _libs.tslibs 模块
hiddenimports += [
    # 核心数据结构（DataFrame 等）
    'pandas.core',
    'pandas.core.frame',  # DataFrame 所在模块
    'pandas.core.dtypes',  # 数据类型相关（to_numeric 依赖）
    'pandas.core.tools',   # to_numeric、to_datetime 工具函数
    # Excel 读写相关（read_excel、ExcelFile 依赖）
    'pandas.io',
    'pandas.io.excel',
    'pandas.io.excel._openpyxl',  # 显式指定 openpyxl 引擎（如果用 openpyxl 读 Excel）
    # 底层计算库（isna、时间处理等依赖）
    'pandas._libs',
    'pandas._libs.tslibs',  # 时间处理（to_datetime 依赖）
    'pandas._libs.missing',  # 缺失值处理（isna 依赖）
    # 其他可能的辅助模块
    'pandas.compat',  # 兼容性处理
    'pandas.util',    # 工具函数
]

# 补充 pandas 所需的数据文件（如配置、模板等）
datas += collect_data_files('pandas')
