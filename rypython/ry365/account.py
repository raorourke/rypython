from __future__ import annotations

import logging
import os
import re
import sys
from O365 import Account, FileSystemTokenBackend
from O365.connection import MSGraphProtocol
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qs, urlparse, unquote

from rypython.ry365.sharepoint import Sharepoint

logging.basicConfig(level=os.environ.get('LOGLEVEL', 'WARNING'))

DOMAIN = 'welocalize.sharepoint.com'
TOKEN_PATH = Path(os.environ.get('welo365_token_path')).expanduser()
CREDS = (os.environ.get('welo365_client_id'), os.environ.get('welo365_client_secret'))


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

    def get_folder_by_url(self, url: str, document_library_transform: Callable = None):
        # url = unquote(link)
        site, lib_name, *folder_path = unquote(link).split('/')[4:]
        account = O365Account(site=site)
        # lib_name = lib_name.replace('Lex 30', 'Lex 3.0')
        if document_library_transform:
            lib_name = document_library_transform(lib_name)
        folder_path = '/'.join(folder_path)
        doc_library = self.get_document_library_by_name(
            lib_name
        ) if lib_name != 'Shared Documents' else site.site.get_default_document_library()
        return doc_library.get_item_by_path(folder_path)

    def get_document_library_by_name(self, document_library_name: str, site: str = None):
        site = self.get_site(site) if site else self.site
        for drive in site.list_document_libraries():
            if drive.name.lower() == document_library_name.lower():
                return drive
