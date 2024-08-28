import pandas as pd
import datetime as dt



#class Reoccurring_Analysis:
#    def __init__(self, object):
df = pd.DataFrame(pd.read_csv('C:/Users/Dispatch2/Desktop/DataAutomation/reoccurring/reoccurring_transports.xlsx', skiprows=1))
date_range = pd.DataFrame(pd.read_csv('C:/Users/Dispatch2/Desktop/DataAutomation/reoccurring/reoccurring_transports.xlsx', nrows=1))

df_no_nans = df.dropna(axis=0) 
df_unique_only = df_no_nans.drop_duplicates()
df_unique_only.rename(columns={'Passenger\'s First Name': 'first', 'Passenger\'s Last Name': 'last', 'Standing Order Start Date': 'start', 'Standing Order End Date': 'end'}, inplace=True)
df_unique_only['end'] = pd.to_datetime(df_unique_only['end']).apply(lambda x: x.date())
#df_unique_only['end'] = df_unique_only['end'].apply(dt.date.)

df_unique_only.reset_index(inplace=True, drop=True)
print(df_unique_only.head(20))
current_date = dt.date.today()

#for row in df_unique_only:
#    if current_date - df_unique_only['end'] < pd.Timedelta('30 days'):
#        print(df_unique_only.iloc[row][['first', 'last', 'end']])
#    else:
#        pass
#print(df_unique_only.iloc[6]['end'] - current_date)
    
        

#for i in range(1, len(df_unique_only)):
#    if df_unique_only.iloc[i]['end'] - current_date < pd.Timedelta('31 days'):
#        print(df_unique_only[i][['first', 'last', 'end']])
#    else:
#        pass

#print(df_unique_only.head(20))
        