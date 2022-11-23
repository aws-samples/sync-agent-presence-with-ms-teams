# Copyright Amazon.com, Inc. or its affiliates. All Rights Reserved.
#SPDX-License-Identifier: MIT-0
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this
# software and associated documentation files (the "Software"), to deal in the Software
# without restriction, including without limitation the rights to use, copy, modify,
# merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
# permit persons to whom the Software is furnished to do so.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
# INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
# PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
# OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
# SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
import os
import json
import logging
import boto3
from requests import get
from requests import post

logger = logging.getLogger()
logger.setLevel(logging.INFO)

secret = os.environ['Secret']

def get_o365_token_header():
    try:
        global app_id 
        client = boto3.client('secretsmanager')
        response = client.get_secret_value( SecretId=secret)
        secrets = json.loads(response['SecretString'])
        
        app_id = secrets['app_id'] 
        client_secret = secrets['client_secret']
        tenant_id = secrets['tenant_id']
        token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/token'
        
        token_data_service_app = {
            'grant_type': 'client_credentials',
            'client_id': app_id,
            'client_secret': client_secret,
            'resource': 'https://graph.microsoft.com',
            'scope': 'https://graph.microsoft.com/.default'
        }
        
        token_r = post(token_url, data=token_data_service_app)
        token = token_r.json().get('access_token')
        headers = {
            'Authorization': 'Bearer {}'.format(token),
            'Content-Type': 'application/json'
        }
        
        return headers 

    except Exception as ex:
        logger.error(ex)


def o365_set_user_presence(o365_id):
    try:
        header = get_o365_token_header()
        logger.info("Updating o365 presence for user: " + str(o365_id))
        endpoint = f'https://graph.microsoft.com/v1.0/users/{o365_id}/presence/setPresence'
        
        # setting expirationDuration to 1 hour
        # this could be changed based on 
        # your needs by updating expirationDuration
        
        user_status = {
            "sessionId": app_id,
            "availability": "Busy",
            "activity": "InACall",
            "expirationDuration": "PT1H"
        }
        response = post(
            url=endpoint,
            headers= header,
            data=json.dumps(user_status)
        )
        logger.info("o365 presence set to busy, response status code: " + str(response.status_code))

    except Exception as ex:
        logger.error(ex)


def o365_clear_user_presence(o365_id):
    try:
        logger.info("Updating o365 presence for user: " + str(o365_id))
        endpoint = f'https://graph.microsoft.com/v1.0/users/{o365_id}/presence/clearPresence'
        status = {
            "sessionId": app_id
        }
        response = post(
            url=endpoint,
            headers=get_o365_token_header(),
            data=json.dumps(status)
        )
        logger.info("o365 presence cleared, response status code: " + str(response.status_code))

    except Exception as ex:
        logger.error(ex)


def o365_get_user_by_email(user):
    try:
        logger.info("Getting o365 user id using email: " + user)
        endpoint = 'https://graph.microsoft.com/v1.0/users/' + user

        response = get(
            url=endpoint,
            headers=get_o365_token_header()
        )
        user = response.json().get('id')
        logger.info("o365 user id: " + user)
        return user

    except Exception as ex:
        logger.error(ex)


def lambda_handler(event, context):
    try:
        action = event['Details']['Parameters']['action']
        agent = event['Details']['Parameters']['agent']
        logger.info("Amazon Connect agent : " + agent)
        logger.info("Amazon Connect action : " + action)

        if action == "agent_busy":
            o365_set_user_presence(o365_get_user_by_email(agent))

        elif action == "agent_free":
            o365_clear_user_presence(o365_get_user_by_email(agent))
            
        elif action != "agent_free" or "agent_busy":
            logger.error("unknown action")
       
    except Exception as ex:
        logger.error(ex)
