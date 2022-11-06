from __future__ import print_function

import logging
import os.path
from typing import Any, Optional, List
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dataclasses import dataclass


@dataclass
class GoogleClient:
    scopes: List[str]
    service: Optional[Any] = None
    creds: Optional[Any] = None

    def __post_init__(self):
        self.creds = None
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file(
                'token.json', self.scopes)

        # If there are no (valid) credentials available, let the user log in.
        if not  self.creds or not  self.creds.valid:
            if  self.creds and  self.creds.expired and  self.creds.refresh_token:
                 self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.scopes)
                self.creds = flow.run_local_server(port=3000)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())

    @property
    def drive_service(self):
        try:
            return build('drive', 'v3', credentials=self.creds)
        except HttpError as err:
            logging.error(err)
            raise err

    @property
    def slide_service(self):
        try:
            return build('slides', 'v1', credentials=self.creds)
        except HttpError as err:
            logging.error(err)
            raise err
