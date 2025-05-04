from setuptools import setup

APP = ['invoice_run.py']  # 메인 파이썬 파일명
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'includes': [
        'cmath', 'datetime', 'pytz',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timezones'
    ],
    'packages': ['pandas', 'openpyxl'],
    'plist': {
        'CFBundleName': '송장파일변환기',
        'CFBundleDisplayName': '송장파일변환기',
        'CFBundleIdentifier': 'com.midnightaxi.invoice_converter',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
    },
    #'iconfile': 'icon.icns',  # 아이콘 파일 필요 시 사용, 없으면 제거 가능
}

setup(
    app=APP,
    name='송장파일변환기',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
