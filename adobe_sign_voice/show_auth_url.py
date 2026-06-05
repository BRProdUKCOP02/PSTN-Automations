import os
from dotenv import load_dotenv
from pathlib import Path
load_dotenv(Path('.env'))
CLIENT_ID = os.getenv('ADOBE_SIGN_CLIENT_ID', '')
REDIRECT_URI = os.getenv('ADOBE_SIGN_REDIRECT_URI', '')
SCOPES = '+'.join(['agreement_read','agreement_write','agreement_send','library_read','user_read'])
auth_url = (
    'https://secure.echosign.com/public/oauth/v2'
    '?response_type=code'
    '&client_id=' + CLIENT_ID
    + '&redirect_uri=' + REDIRECT_URI
    + '&scope=' + SCOPES
)
print('CLIENT_ID   :', CLIENT_ID)
print('REDIRECT_URI:', REDIRECT_URI)
print()
print('Full auth URL:')
print(auth_url)
