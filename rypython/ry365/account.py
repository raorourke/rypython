from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from urllib.parse import parse_qs, urlparse, unquote

import yaml
from O365 import Account, FileSystemTokenBackend
from O365.connection import MSGraphProtocol

from rypython.ry365.sharepoint import Sharepoint

logging.basicConfig(level=os.environ.get('LOGLEVEL', 'WARNING'))

RYPYTHON_CONFIG_PATH = os.environ.get(
    'RYPYTHON_CONFIG_PATH',
    Path.home() / '.config' / 'rypython'
)
if isinstance(RYPYTHON_CONFIG_PATH, str):
    RYPYTHON_CONFIG_PATH = Path(RYPYTHON_CONFIG_PATH)

if not RYPYTHON_CONFIG_PATH:
    raise MissingConfiguration('Missing local configuration path!')
RYPYTHON_CONFIG_FILE = RYPYTHON_CONFIG_PATH / 'config.yaml'
if not RYPYTHON_CONFIG_FILE:
    raise MissingConfiguration('Missing local configuration file!')

with open(RYPYTHON_CONFIG_FILE, 'r') as configfile:
    RYPYTHON_CONFIG = yaml.safe_load(configfile)

DOMAIN = RYPYTHON_CONFIG['sharepoint']['domain']
CLIENT_ID = RYPYTHON_CONFIG['sharepoint']['client_id']
CLIENT_SECRET = RYPYTHON_CONFIG['sharepoint']['client_secret']
CREDS = (CLIENT_ID, CLIENT_SECRET)
TOKEN_PATH = os.environ.get(
    'RYPYTHON_TOKEN_PATH',
    Path.home() / '.tokens'
)
if isinstance(TOKEN_PATH, str):
    TOKEN_PATH = Path(TOKEN_PATH)

'''
class TokenBackend(FileSystemTokenBackend):
    def save_token(self):
        """
        Saves the token dict in the specified file
        :return bool: Success / Failure
        """
        if self.token is None:
            raise ValueError('You have to set the "token" first.')

        try:
            if not self.token_path.parent.exists():
                self.token_path.parent.mkdir(parents=True)
        except Exception as e:
            logging.error('Token could not be saved: {}'.format(str(e)))
            return False

        with self.token_path.open('w') as token_file:
            # 'indent = True' will make the file human readable
            self.serializer.dump(self.token, token_file, indent=True)

        return True

'''


class O365Account(Account):
    def __init__(
            self,
            site: str = None,
            creds: tuple[str, str] = CREDS,
            scopes: list[str] = None,
            auth_flow_type: str = 'authorization'
    ):

        TOKEN = TOKEN_PATH / 'o365_token.txt'
        token_backend = None
        if TOKEN.exists():
            logging.debug(f"Using token file {TOKEN}")
            token_backend = FileSystemTokenBackend(token_path=TOKEN_PATH)
            token_backend.load_token()
            token_backend.get_token()
        scopes = scopes or ['offline_access', 'Sites.Manage.All']
        OPTIONS = {
            'token_backend': token_backend
        } if token_backend is not None else {
            'scopes': scopes,
            'auth_flow_type': auth_flow_type
        }
        super().__init__(creds, **OPTIONS)
        if not self.is_authenticated:
            self.authenticate()
        self.drives = self.storage().get_drives()
        self.site = self.get_site(site) if site else None
        self.drive = self.site.get_default_document_library() if self.site else self.storage().get_default_drive()
        self.root_folder = self.drive.get_root_folder()

    def sharepoint(self, *, resource=''):
        if not isinstance(self.protocol, MSGraphProtocol):
            raise RuntimeError(
                'Sharepoint api only works on Microsoft Graph API'
            )
        return Sharepoint(parent=self, main_resource=resource)

    def get_site(self, site: str):
        return self.sharepoint().get_site(DOMAIN, f"/sites/{site}")

    def search(self, query: str):
        u = urlparse(query)
        site = None
        drive = None
        file_name = None
        if (site_query := re.search(r'.*/sites/(?P<site>[\w_\-\d]+)/', u.path)):
            if (site := site_query.group('site')):
                site = self.get_site(site)
                drive = site.get_default_document_library()
        if (q := parse_qs(u.query)):
            file_name = q.get('file', [''])[0]
        drive = drive or self.drive
        query = file_name or query
        results = list(drive.search(query))
        return results[0] if results else results

    def get_folder(self, *subfolders: str, site: str = None):
        if len(subfolders) == 0:
            return self.drive

        site = self.get_site(site) if site else self.site
        drive = site.get_default_document_library() if site else self.drive

        items = drive.get_items()
        for subfolder in subfolders:
            try:
                subfolder_drive = list(filter(lambda x: subfolder in x.name, items))[0]
                items = subfolder_drive.get_items()
            except:
                raise ('Path {} not exist.'.format('/'.join(subfolders)))
        return subfolder_drive

    @staticmethod
    def get_item_by_url(item_path: str, library_map: dict = None):
        pattern = r'^.*\/sites\/(?P<site>[0-9.\-A-Za-z]+)\/(?:[0-9.\-A-Za-z\s]+)\/(?P<folder_path>.*)$'
        site, lib_name, *item_path = unquote(item_path).split('/')[4:]
        item_path = '/'.join(item_path)
        account = O365Account(site=site)
        doc_library = None
        if lib_name == 'Shared Documents':
            doc_library = account.site.get_default_document_library()
        if lib_name != 'Shared Documents':
            if library_map:
                lib_name = library_map.get(lib_name, lib_name)
            doc_library = account.get_document_library_by_name(lib_name)
        if not doc_library:
            raise ValueError(f"Could not find document library named {lib_name} on site {site}!")
        return doc_library.get_item_by_path(item_path)

    def get_folder_by_path(self, folder_path: str):
        pattern = r'^.*\/sites\/(?P<site>[0-9.\-A-Za-z]+)\/(?:[0-9.\-A-Za-z\s]+)\/(?P<folder_path>.*)$'
        folder_path = unquote(folder_path)
        matches = re.search(pattern, folder_path)
        if not matches:
            raise ValueError(f"Folder not found. Check that your path.")
        matches = matches.groupdict()
        site = matches.get('site')
        folder_path = matches.get('folder_path')
        if not site or not folder_path:
            raise ValueError(f"Folder not found. Check that your path.")
        return self.get_folder(
            *folder_path.split('/'),
            site=site
        )

    def get_folder_by_url(self, url: str, library_map: dict = None):
        site, lib_name, *folder_path = unquote(url).split('/')[4:]
        folder_path = '/'.join(folder_path)
        if site != self.site.name:
            return O365Account(site=site).get_folder_by_url(
                url, library_map=library_map
            )
        if lib_name != 'Shared Documents' and library_map:
            lib_name = library_map.get(lib_name)
        doc_library = self.get_document_library_by_name(
            lib_name
        ) if lib_name != 'Shared Documents' else self.site.get_default_document_library()
        return doc_library.get_item_by_path(folder_path)

    def get_document_library_by_name(self, document_library_name: str, site: str = None):
        site = self.get_site(site) if site else self.site
        for drive in site.list_document_libraries():
            if drive.name.lower() == document_library_name.lower():
                return drive
