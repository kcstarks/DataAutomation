import asyncio
import configparser
import os
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph
import base64
from analysis import Analyze
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Alignment 
import shutil
import datetime
from cleaner import cleaner


async def main():
    print('...waiting on request')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    graph: Graph = Graph(azure_settings)

    cleaner()

    today = datetime.datetime.today()
    weekday = today.weekday()
    day = today.day

    try:
        await create_daily_report(graph)

        if weekday != 0:
            pass
        else:
            await create_weekly_report(graph)

        if day != 1:
            pass
        else:
            await create_monthly_report(graph)

    except ODataError as odata_error:
        print('Error:')
        if odata_error.error:
            print(odata_error.error.code, odata_error.error.message)
    



#-------------------Graph Functions-----------------------------------
    
async def get_inboxes(graph: Graph, folderstr: str):
    folders_page = await graph.get_inbox(folderstr)
    if folders_page and folders_page.value:
        for folder in folders_page.value:
            print('Folder Name: ', folder.display_name, 'ID: ', folder.id)

async def get_message_id(graph: Graph, folder: str):
    emails_page = await graph.get_message_id(folder)
    if emails_page and emails_page.value:
      for email in emails_page.value:
        return email.id

async def get_attachment_id(graph: Graph, message_id, folder: str):
    attachments = await graph.get_attachment_id(message_id, folder)
    if attachments and attachments.value:
         for attachment in attachments.value:
            return attachment.id

async def get_attachment(graph: Graph, message_id, attachment_id, folder):
    #might need to be separated into two functions for function purity
    attachment = await graph.get_attachment_file(message_id, attachment_id, folder)
    byte = attachment.content_bytes
    byte_decoded = base64.b64decode(byte)
    csv = str(byte_decoded)[2:-1]
    csv_modified = csv.replace('\\r\\n', '\n').replace('\\n', '\n')
    return csv_modified

async def send_email_with_attachment(graph: Graph, date, report_type, byte, file_name):
    print('Sending email...')
    email_list = ['kstarks@ridensafe.com', 'cwise@ridensafe.com']
    for email in email_list:
        await graph.send_email(date=date, report_type=report_type, byte=byte, file_name=file_name, email_address=email)

#-------------------Create Instance Functions-------------------------

async def create_daily_report(graph: Graph):
    daily_report = Daily_Report(graph)
    await daily_report._init(graph)
    return daily_report

async def create_weekly_report(graph: Graph):
    weekly_report = Weekly_Report(graph)
    await weekly_report._init(graph)
    return weekly_report

async def create_monthly_report(graph: Graph):
    monthly_report = Monthly_Report(graph)
    await monthly_report._init(graph)
    return monthly_report


#------------------Analysis Functions--------------------------------

def export_excel(file_path, object):
    template = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/blank.xlsx'

    shutil.copyfile(template, file_path)

    with pd.ExcelWriter(
        file_path,
        mode='a',
        engine='openpyxl',
        if_sheet_exists='overlay'
    ) as writer:
        object.driver_table.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=0, index=False)
        object.mode_table.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=8, index=False)
        object.status_table.to_excel(writer, sheet_name='Sheet1', startrow=7, startcol=8, index=False)

    wb = load_workbook(file_path)
    ws = wb.active
    
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter


        for cell in column:
            try: 
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                print('export_excel() error: cell width error')

        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

        for cell in column:
            try:
                cell.alignment = Alignment(horizontal='center')
            except:
                print('export_excel() error: cell alignment error')
    
    
    wb.save(file_path)
    print('Data exported to excel file')
   


def file_to_byte(file_path):
    with open(file_path, 'rb') as upload:
        media_content = base64.b64encode(upload.read())
        byte = media_content.decode('utf-8')
    return byte


class UploadedFiles:

    def __init__(self, x, y, z) -> None:
        self.drivers = x
        self.orders = y
        self.samsara = z

#-------------------Report Classes-----------------------------------

class Daily_Report:

    def __init__(self, graph: Graph):
        print('initializing daily report')
        #Is this being used?
        self.daily_files_object = UploadedFiles("","","")

        self.driver_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/drivers/driver_report.csv'
        self.order_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/orders/order_report.csv'
        self.samsara_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/samsara/safety_report.csv'
        self.excel_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/daily_report.xlsx'

        self.date_range = ''
        self.report_type = 'Daily Report'
        self.file_name = os.path.basename(self.excel_file)
        
    async def _init(self, graph: Graph):
        self.driver = await self.get_daily_report_drivers(graph)  
        self.order = await self.get_daily_report_orders(graph)
        self.samsara = await self.get_samsara_daily_report(graph)
        print('Analysis initiated...')
        
        analysis_object = UploadedFiles(self.driver_file, self.order_file, self.samsara_file)
        analysis_results = Analyze(analysis_object)
        table_object = analysis_results.data 
        self.date_range = table_object.date_range
            #attributes: date_range, driver_table, mode_table, status_table

        export_excel(self.excel_file, table_object)
        byte = file_to_byte(self.excel_file)

        await send_email_with_attachment(graph, date=self.date_range, report_type=self.report_type, byte=byte, file_name=self.file_name)



    async def get_daily_report_drivers(self, graph: Graph):
        print('pulling daily report for drivers...')
        folder = 'daily_driver'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)

        print(csv, file=open(self.driver_file, 'w', newline='\n'))

    async def get_daily_report_orders(self, graph: Graph):
        print('pulling daily report for orders...')
        folder = 'daily_order'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)

        print(csv, file=open(self.order_file, 'w', newline='\n'))

    async def get_samsara_daily_report(self, graph: Graph):
        print('pulling samsara report...')
        folder = 'samsara_daily'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)

        print(csv, file=open(self.samsara_file, 'w', newline='\n') )

class Weekly_Report:

    def __init__(self, graph: Graph):
        print('initializing weekly report')

        self.weekly_files_object = UploadedFiles("","","")

        self.driver_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/drivers/driver_report.csv'
        self.order_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/orders/order_report.csv'
        self.samsara_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/samsara/safety_report.csv'
        self.excel_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/weekly_report.xlsx'

        self.date_range = ''
        self.report_type = 'Weekly Report'
        self.file_name = os.path.basename(self.excel_file)

    async def _init(self, graph: Graph):
        self.driver = await self.get_weekly_report_drivers(graph)  
        self.order = await self.get_weekly_report_orders(graph)
        self.samsara = await self.get_samsara_weekly_report(graph)

        print('Analysis initiated...')
        analysis_object = UploadedFiles(self.driver_file, self.order_file, self.samsara_file)
        analysis_results = Analyze(analysis_object)
        table_object = analysis_results.data 
        self.date_range = table_object.date_range
            #attributes: date_range, driver_table, mode_table, status_table
        export_excel(self.excel_file, table_object)
        byte = file_to_byte(self.excel_file)

        await send_email_with_attachment(graph, date=self.date_range, report_type=self.report_type, byte=byte, file_name=self.file_name)

    async def get_weekly_report_drivers(self, graph: Graph):
        print('pulling weekly report for drivers...')
        folder = 'weekly_driver'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)

        print(csv, file=open(self.driver_file, 'w', newline='\n'))

    async def get_weekly_report_orders(self, graph: Graph):
        print('pulling weekly report for orders...')
        folder = 'weekly_order'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)

        print(csv, file=open(self.order_file, 'w', newline='\n'))

    async def get_samsara_weekly_report(self, graph: Graph):
        print('pulling samsara report...')
        folder = 'samsara_weekly'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)
        print(csv, file=open(self.samsara_file, 'w', newline='\n') )




class Monthly_Report:

    def __init__(self, graph: Graph):
        print('initializing monthly report')

        self.monthly_files_object = UploadedFiles("","","")

        self.driver_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/drivers/driver_report.csv'
        self.order_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/orders/order_report.csv'
        self.samsara_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/samsara/safety_report.csv'
        self.excel_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/monthly_report.xlsx'

        self.date_range = ''
        self.report_type = 'Monthly Report'
        self.file_name = os.path.basename(self.excel_file)

    async def _init(self, graph: Graph):
        self.driver = await self.get_monthly_report_drivers(graph)  
        self.order = await self.get_monthly_report_orders(graph)
        self.samsara = await self.get_samsara_monthly_report(graph)
        print('Analysis initiated...')

        analysis_object = UploadedFiles(self.driver_file, self.order_file, self.samsara_file)
        analysis_results = Analyze(analysis_object)
        table_object = analysis_results.data 
        self.date_range = table_object.date_range
            #attributes: date_range, driver_table, mode_table, status_table
        export_excel(self.excel_file, table_object)
        byte = file_to_byte(self.excel_file)

        await send_email_with_attachment(graph, date=self.date_range, report_type=self.report_type, byte=byte, file_name=self.file_name)

    async def get_monthly_report_drivers(self, graph: Graph):
        print('pulling monthly report for drivers...')
        folder = 'monthly_driver'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)

        print(csv, file=open(self.driver_file, 'w', newline='\n'))

    async def get_monthly_report_orders(self, graph: Graph):
        print('pulling monthly report for orders...')
        folder = 'monthly_order'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)

        print(csv, file=open(self.order_file, 'w', newline='\n'))
 
    async def get_samsara_monthly_report(self, graph: Graph):
        print('pulling samsara report...')
        folder = 'samsara_monthly'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)
        print(csv, file=open(self.samsara_file, 'w', newline='\n') )




if __name__ == '__main__':
    asyncio.run(main())
   