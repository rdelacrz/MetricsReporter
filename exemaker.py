"""
This module is responsible for turning the run script into a runnable 
executable.
"""

from distutils.core import setup
import py2exe, sys

sys.argv.append('py2exe')

setup(
    options = {
               'py2exe' : 
                    {
                        'dll_excludes' : [ "mswsock.dll", "powrprof.dll", "OCI.dll" ],
                        'includes' : ["decimal"],
                        'bundle_files' : 1
                    }
               },
    windows = [{'script' : "run.py"}],
    zipfile = None
)