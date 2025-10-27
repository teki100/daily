from PyInstaller.utils.hooks import collect_all, get_module_file_attribute
import os

# 获取pandas的安装路径（确保该路径在当前环境中存在）
pandas_path = os.path.dirname(get_module_file_attribute('pandas'))
# 收集该路径下的所有内容
datas, binaries, hiddenimports = collect_all('pandas', src=pandas_path)
# 补充核心子模块（防止遗漏）
hiddenimports += [
    'pandas',
    'pandas._libs',
    'pandas._libs.tslibs',
    'pandas.core',
    'pandas.io',
]
