import os 

def cleaner():

    print('Cleaning...')

    folder_dictionary = [
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/drivers/driver_report.csv',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/orders/order_report.csv',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/daily_report.xlsx',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/samsara/safety_report.csv'


        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/drivers/driver_report.csv',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/orders/order_report.csv',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/weekly_report.xlsx',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/weekly/samsara/safety_report.csv'

        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/drivers/driver_report.csv',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/orders/order_report.csv',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/monthly_report.csv',
        'C:/Users/Dispatch2/Desktop/DataAutomation/reports/monthly/samsara/safety_report.csv'
    ]
   
    for file in folder_dictionary:
        if os.path.exists(file): 
            os.remove(file)
        else:
            return
        