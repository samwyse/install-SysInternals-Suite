'''downloads and installs the SysInternals suite

This program downloads the SysInternals suite ZIP file, expands it
into the appropriate Program Files location, and creates a Start
menu entry containing all of the GUI programs.'''

from argparse import ArgumentParser
from os import environ, makedirs, unlink, walk
from os.path import isdir, join
from platform import python_version_tuple
from requests import get
from shutil import rmtree
from StringIO import StringIO
from zipfile import ZipFile

from logging import basicConfig, error, info, INFO
import pefile
import win32com.client

class Installer(object):
        
    def __init__(self, **kwds):
        self.__dict__.update(kwds)
        
    def download_url(self):
        if self.uninstall:
            return
        info('downloading %s', self.url)
        response = get(self.url)
        assert response.status_code == 200
        assert self.content_type in response.headers['content-type']
        self.content = response.content

    def _mksubdir(self, path, name, uninstall):
        assert isdir(path)
        path = join(path, name)
        if uninstall:
            rmtree(path, ignore_errors=True)
        elif not isdir(path):
            makedirs(path)
        else:
            pass
        return path

    def process_subdirectories(self):
        info('%s subdirectories' % ('removing' if self.uninstall else 'creating'))
        if self.allusers:
            Programs = sh.SpecialFolders('AllUsersPrograms')
            ProgramFiles = environ['ProgramFiles']
        else:
            Programs = sh.SpecialFolders('Programs')
            ProgramFiles = environ['LOCALAPPDATA']

        self.Programs = self._mksubdir(Programs, self.group, self.uninstall)
        self.ProgramFiles = self._mksubdir(ProgramFiles, self.group, self.uninstall)

    def extract_programs(self):
        if self.uninstall:
            return
        info('extracting programs')
        ZipFile(StringIO(self.content)).extractall(self.ProgramFiles)

    def create_links(self):
        if self.uninstall:
            return
        info('creating links')
        for root, dirs, files in walk(self.ProgramFiles):
            for name in files:
                if name.lower().endswith('.exe'):
                    shortcut_name = name[:-4]
                    shortcut_target = join(root, name)
                    pe = pefile.PE(shortcut_target)
                    if pe.OPTIONAL_HEADER.Subsystem == gui_app:
                        lnk = sh.CreateShortcut(join(self.Programs, shortcut_name + '.lnk')) 
                        lnk.TargetPath = shortcut_target
                        lnk.Save()

    def run(self):
        self.download_url()
        self.process_subdirectories()
        self.extract_programs()
        self.create_links()

# make sure we have a safe ZipFile.extractall
assert python_version_tuple() >= (2, 7, 4)

sh = win32com.client.Dispatch('WScript.Shell')
pefile.fast_load = True
gui_app = pefile.SUBSYSTEM_TYPE['IMAGE_SUBSYSTEM_WINDOWS_GUI']

if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.set_defaults(
        group = 'SysInternals',
        url = 'http://download.sysinternals.com/files/SysinternalsSuite.zip',
        content_type = 'zip',
        )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-i', '--install', help='install the application for the current user', action='store_true')
    group.add_argument('-a', '--allusers', help='install the application for all users', action='store_true')
    parser.add_argument('-x', '--uninstall', help='uninstall the application', action='store_true')
    parser.add_argument('-v', '--verbose', help='set output verbosity', action='store_true')
    parser.add_argument('--group', help='specify shortcut group', type=str)
    parser.add_argument('--url', help='specify url to download', type=str)
    parser.parse_args(namespace=Installer()).run()
