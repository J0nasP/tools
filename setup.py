from distutils.core import setup
import py2exe


setup(
        windows=['Toolbox.py'],
        options={
                "py2exe":{
                        "unbuffered": True,
                        "optimize": 2,
                        "includes": ['lxml.etree', 'lxml._elementpath', 'gzip', 'docx2pdf']
                }
        }
)