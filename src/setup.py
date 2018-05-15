import os
import sys
from os.path import join, dirname, abspath

from distutils.core import setup
import py2exe
import matplotlib
setup(
    windows=[
        {
         "script":'excel_to_python.py',
         #"icon_resources":[(0, "icon.ico")]
         }
        ],
    data_files=matplotlib.get_py2exe_datafiles(),
    options={'py2exe': {
            "dist_dir": join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))),'dist'),
            "includes" : ["matplotlib.backends.backend_wxagg"],
            'excludes': ['_gtkagg','_tkagg'],
            #'dll_excludes' : ["libopenblas.BNVRK7633HSX7YVO2TADGR4A5KEKXJAW.gfortran-win_amd64.dll"]
            }}
)
