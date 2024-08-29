#This file is for alerting users to reoccurring transports that are expiring soon.

import configparser
import shutil
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
import asyncio

from openpyxl import load_workbook
from openpyxl.styles import Alignment 
import pandas as pd
from graph import Graph
from main import get_attachment, get_attachment_id, get_inboxes, get_message_id
from pd_reoccurring import Reoccurring_Analysis


async def reoccurring():
    print('Gathering reoccurring transports...')

    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    graph: Graph = Graph(azure_settings)


    try:
        print('testing')
        await create_reoccurring(graph)
    except ODataError as odata_error:
        print('Error:')
        if odata_error.error:
            print(odata_error.error.code, odata_error.error.message)


#------------------------------------------------------

async def create_reoccurring(graph):
    reoccurring = Reoccurring_Transports(graph)
    await reoccurring._init(graph)
    return reoccurring

def export_excel(file_path, table):
    template = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/blank.xlsx'

    shutil.copyfile(template, file_path)

    with pd.ExcelWriter(
        file_path,
        mode='a',
        engine='openpyxl',
        if_sheet_exists='overlay'
    ) as writer:
        table.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=0, index=False)


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
   



class Reoccurring_Transports:
    
    def __init__(self, graph:Graph):
        self.reoccurring_file_csv = 'C:/Users/Dispatch2/Desktop/DataAutomation/reoccurring/reoccurring_transports.csv'
        self.reoccurring_file_excel = 'C:/Users/Dispatch2/Desktop/DataAutomation/reoccurring/reoccurring_transports.xlsx'

    async def _init(self, graph:Graph):
        print('...')
        await self.mk_reoccurring_csv(graph)
        analysis_result = Reoccurring_Analysis()
        data = analysis_result.data
        table = data.table
        date = data.date_range

        export_excel(self.reoccurring_file_excel, table)

    async def mk_reoccurring_csv(self, graph:Graph):
        folder = 'reoccurring'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)
        
        print(csv, file=open(self.reoccurring_file_csv, 'w', newline='\n'))

 
asyncio.run(reoccurring())