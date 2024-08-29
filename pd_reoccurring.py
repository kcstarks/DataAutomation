import pandas as pd
import datetime as dt

class ReturnObject:
    def __init__(self, table):
        self.table = table

class Reoccurring_Analysis:
    def __init__(self):
        df = pd.DataFrame(pd.read_csv('C:/Users/Dispatch2/Desktop/DataAutomation/reoccurring/reoccurring_transports.csv', skiprows=1))

        df_no_nans = df.dropna(axis=0) 
        df_unique_only = df_no_nans.drop_duplicates()
        df_unique_only = df_unique_only.rename(columns={'Passenger\'s First Name': 'first', 'Passenger\'s Last Name': 'last', 'Standing Order Start Date': 'start', 'Standing Order End Date': 'end'})
        df_unique_only.reset_index(drop=True, inplace=True)
        df_unique_only['end'] = pd.to_datetime(df_unique_only['end'], format='%Y-%m-%d')
        df_unique_only['end'] = df_unique_only['end'].apply(lambda x: x.date())

        current_date = dt.date.today()

        df_ending_soon = pd.DataFrame(columns=['First Name', 'Last Name', 'End Date'])

        for i in range(1, len(df_unique_only)):
            if df_unique_only.iloc[i]['end'] - current_date > pd.Timedelta('31 days'):
                pass
            elif df_unique_only.iloc[i]['end'] - current_date < pd.Timedelta('0 days'):
                pass
            else:
                row_to_add = pd.DataFrame(
                    [{
                        'First Name': df_unique_only.iloc[i]['first'], 
                        'Last Name': df_unique_only.iloc[i]['last'], 
                        'End Date': df_unique_only.iloc[i]['end']
                    }]
                )
                df_ending_soon = pd.concat([df_ending_soon, row_to_add])

        df_ending_soon.reset_index(drop=True, inplace=True)
        
        self.data = ReturnObject(
            table = df_ending_soon, 
        )