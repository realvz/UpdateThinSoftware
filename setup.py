from distutils.core import setup
import py2exe
setup(
        console=['__init__.py'],
        #zipfile = None,
        options={
                "py2exe":{
                        "unbuffered": True,
                        "optimize": 2,
                        "includes": ["email"]
                        #"bundle_files": 1,
                }
        })