from distutils.core import setup
import importlib.metadata
import py2exe


setup(
        windows=['Toolbox.py'],
        options={
                "py2exe":{
                        "optimize": 2,
                        "includes": ['lxml.etree', 'lxml._elementpath', 'gzip', 'tqdm', "win32api", "System", "_posixshmem", "_uuid", "collections.MutableMapping", "collections.Sequence", "colorama", "importlib_metadata", "ipywidgets", "matplotlib", "multiprocessing.RLock", "ordereddict" "pandas", "pkg_resources.extern.appdirs", "pkg_resources.extern.packaging", "pkg_resources.extern.six", "readline", "resource", "sets"],
                        "compressed": 2,
                        "unbuffered": 2,
                        "packages": ["docx2pdf"],

                }

        }
)