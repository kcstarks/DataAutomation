import base64
from configparser import SectionProxy
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder)
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody)
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.attachment import Attachment
from msgraph.generated.models.file_attachment import FileAttachment

class Graph:
    settings: SectionProxy
    client_credential: ClientSecretCredential
    app_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        client_secret = self.settings['clientSecret']

        self.client_credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        self.scopes = ['https://graph.microsoft.com/.default']
        self.app_client = GraphServiceClient(credentials=self.client_credential, scopes=self.scopes) 

        self.user = 'deae0de2-15e8-425e-8a5a-ef516e856783'

        #Parent Folder (AppData Folder)
        self.app_data_folder_id = "AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADnQdigAAA="
        self.app_data_folder_path = self.app_client.users.by_user_id(self.user).mail_folders.by_mail_folder_id(self.app_data_folder_id)

        #Child Folders (AppData Child Folders)
        self.daily_reports_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADnQdihAAA='
        self.monthly_reports_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADnQdijAAA='
        self.weekly_reports_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADnQdiiAAA='

        #Child Folder Paths 
        self.daily_reports_folder_path = self.app_client.users.by_user_id(self.user).mail_folders.by_mail_folder_id(self.daily_reports_folder_id)
        self.weekly_reports_folder_path = self.app_client.users.by_user_id(self.user).mail_folders.by_mail_folder_id(self.weekly_reports_folder_id)
        self.daily_reports_folder_path = self.app_client.users.by_user_id(self.user).mail_folders.by_mail_folder_id(self.daily_reports_folder_id)

        #Child Sub Folders
        self.daily_report_drivers_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADr1XP7AAA='
        self.daily_report_orders_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADr1XP8AAA='
        
        self.weekly_report_drivers_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADvJ4xTAAA='
        self.weekly_report_orders_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADvJ4xUAAA='

        self.monthly_report_drivers_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADvJ4xRAAA='
        self.monthly_report_orders_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAADvJ4xSAAA='
    
        self.samsara_daily_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAAD7hZ_yAAA='        
        self.samsara_weekly_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAAD7hZ_zAAA='        
        self.samsara_monthly_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAAD7hZ_0AAA='        

        #Samsara Folder (AppData Child Folder)
        self.samsara_folder_id = 'AAMkAGUxYjNhMjhjLTAwNDgtNDBlMy1iNWJhLWVhMWI2YWY0N2FiNgAuAAAAAABnaMnddrDARatjpjvAYqRzAQCjWYOTFOeTTZkOXwdENkRPAAD7hZ_xAAA='


        self.folder_dictionary = {
            'daily_driver': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.daily_report_drivers_folder_id),
            'daily_order': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.daily_report_orders_folder_id),

            'weekly_driver': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.weekly_report_drivers_folder_id),
            'weekly_order': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.weekly_report_orders_folder_id),

            'monthly_driver': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.monthly_report_drivers_folder_id),
            'monthly_order': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.monthly_report_orders_folder_id),

            'samsara_daily': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.samsara_daily_folder_id),
            'samsara_weekly': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.samsara_weekly_folder_id),
            'samsara_monthly': self.app_data_folder_path.child_folders.by_mail_folder_id1(self.samsara_monthly_folder_id),
        }

        self.folders = {
            'daily': self.daily_reports_folder_id,
            'weekly': self.weekly_reports_folder_id,
            'monthly': self.monthly_reports_folder_id,

            'samsara_daily': self.samsara_daily_folder_id,
            'samsara_weekly': self.samsara_weekly_folder_id,
            'samsara_monthly': self.samsara_monthly_folder_id,
        }



    async def get_app_only_token(self):
        graph_scope = 'https://graph.microsoft.com/.default'
        access_token = await self.client_credential.get_token(graph_scope)
        return access_token.token
    
    async def get_user(self):
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            # Only request specific properties
            select = ['id']
        )
        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        user = await self.app_client.users.by_user_id('kstarks@ridensafe.com').get(request_configuration=request_config)
        return user
    

    async def get_inbox(self, folder:str):
    
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select= ['displayName', 'id'],
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        message = await self.app_client.users.by_user_id(self.user).mail_folders.by_mail_folder_id(self.folders[folder]).child_folders.get(request_configuration=request_config)
        return message
    
    async def get_message_id(self, folder: str):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select= ['id'],
            top=1
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        message_id = await self.folder_dictionary[folder].messages.get(request_configuration=request_config)
        return message_id


    async def get_attachment_id(self, message_id, folder: str):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select = ['id']
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )
        attachemnt_id = await self.folder_dictionary[folder].messages.by_message_id(message_id).attachments.get(request_configuration=request_config)
        return attachemnt_id

    async def get_attachment_file(self, message_id, attachment_id, folder):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            expand= ['microsoft.graph.itemattachment/item']
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )
        attachment_file = await self.folder_dictionary[folder].messages.by_message_id(message_id).attachments.by_attachment_id(attachment_id).get(request_configuration=request_config)
        return attachment_file
    

    async def send_email(self, date, report_type:str, byte:bytes, file_name:str, email_address:str):
        request_body = SendMailPostRequestBody(
        	message = Message(
        		subject = report_type,
        		body = ItemBody(
        			content_type = BodyType.Text,
        			content = '{report} for {date}'.format(report=report_type, date=date)
        		),
                importance= 'normal',
        		to_recipients = [
        			Recipient(
        				email_address = EmailAddress(
        					address = email_address,
        				),
        			),
        		],
        		attachments = [
			        FileAttachment(
			        	odata_type = "#microsoft.graph.fileAttachment",
			        	name = file_name,
			        	content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
			        	content_bytes = base64.urlsafe_b64decode(byte),
			        ),
        		],
        	),
        )
        await self.app_client.users.by_user_id(self.user).send_mail.post(request_body)
        
    
   