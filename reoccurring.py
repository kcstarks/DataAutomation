#This file is for alerting users to reoccurring transports that are expiring soon.

import configparser
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
import asyncio
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

class Upload_File:
    def __init__(self, file:str) -> None:
        self.csv = file


async def create_reoccurring(graph):
    reoccurring = Reoccurring_Transports(graph)
    await reoccurring._init(graph)
    return reoccurring


class Reoccurring_Transports:
    
    def __init__(self, graph:Graph):
        self.reoccurring_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reoccurring/reoccurring_transports.xlsx'

    async def _init(self, graph:Graph):
        print('...')
        await self.mk_reoccurring_csv(graph)
        analysis_obj = Upload_File(self.reoccurring_file)
        analysis_result = Reoccurring_Analysis(analysis_obj)

    async def mk_reoccurring_csv(self, graph:Graph):
        folder = 'reoccurring'
        message_id = await get_message_id(graph, folder)
        attachment_id = await get_attachment_id(graph, message_id, folder)
        csv = await get_attachment(graph, message_id, attachment_id, folder)
        
        print(csv, file=open(self.reoccurring_file, 'w', newline='\n'))

 
asyncio.run(reoccurring())