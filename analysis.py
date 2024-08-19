import pandas as pd
import datetime as dt
import numpy as np

#CLASSES-------------

class ReturnObject:
    def __init__(self, date_range, driver_table, mode_table, status_table):
        self.date_range = date_range
        self.driver_table = driver_table
        self.mode_table = mode_table
        self.status_table = status_table


#FUNCTIONS-----------
#Variables-----------


#MAIN ----
class Analyze(object):
    def __init__(self, object):

        #FILE TO DATAFRAME
        self.date_range = pd.DataFrame(pd.read_csv(object.drivers, nrows=1)).columns.values
        self.df_drivers = pd.DataFrame(pd.read_csv(object.drivers, skiprows=1))
        self.df_orders = pd.DataFrame(pd.read_csv(object.orders, skiprows=1))


        #DATAFRAME CLEANUP
        self.cleanUp(self.df_drivers)
        self.cleanUp(self.df_orders)
        #print(self.df_orders.head(4))
        self.df_drivers.insert(0, "full_name", "")
        self.df_drivers["full_name"] = self.df_drivers["driver_first_name"] + " " + self.df_drivers["driver_last_name"]
        self.df_drivers.drop(columns=['driver_first_name', 'driver_last_name'], inplace=True)
        self.df_drivers[['early_%(per_payer_settings)', 'late_%(per_payer_settings)']] = self.df_drivers[['early_%(per_payer_settings)', 'late_%(per_payer_settings)']].apply(lambda x: round(x, 2))


    
        #NEW DATAFRAMES
        self.orders_driver_price = self.df_orders[["full_name", "order_id", "order_price"]]
        self.orders_by_driver = self.orders_driver_price.groupby(by="full_name").sum()
        #print("orders by driver: ", self.orders_by_driver.head(5))

        self.orders_by_driver_sum = pd.merge(self.df_drivers, self.orders_by_driver, on='full_name')
        self.orders_by_driver_sum.drop(columns='order_id', inplace=True)
        self.orders_by_driver_sum_sorted =self.orders_by_driver_sum.sort_values(by='full_name', ascending=True)
        #self.orders_by_driver_sum_sorted['order_price'].apply(lambda x:'${0:,.0f}'.format(x))
        #print('orders by driver sum sorted: ',self.orders_by_driver_sum_sorted)

        #PIVOT TABLES
        #self.pivot_mode = pd.pivot_table(self.df_orders, index=["order_mode"], values=["order_id"], aggfunc="count").sort_values(by=["order_id"], ascending=False)
        #self.pivot_mode.reset_index(inplace=True)
        #self.pivot_mode = self.pivot_mode.rename(columns= {"order_mode": "Mode","order_id": "Count"})

        def checker(list, dictionary, column, *args):
            for status in args:
                if status in list:
                    dictionary[status] = self.df_orders[column].value_counts()[status]
                else:
                    dictionary[status] = 0

        mode_list = self.df_orders['order_mode'].unique()
        mode_dictionary = {
            'Ambulatory': 0,
            'Wheelchair': 0,
            'Stretcher': 0
        }

        checker(mode_list, mode_dictionary, 'order_mode', 'Ambulatory', 'Wheelchair', 'Stretcher')

        ambulatory = mode_dictionary['Ambulatory']
        wheelchair = mode_dictionary['Wheelchair']
        stretcher = mode_dictionary['Stretcher']
        total_list = [ambulatory, wheelchair, stretcher]
        total = sum(total_list)

        self.mode_table = pd.DataFrame({
            'Mode': ['Ambulatory', 'Wheelchair', 'Stretcher', 'Total'],
            'Count': [ambulatory, wheelchair, stretcher, total],
        })

        #----------------------------------------------------------

        status_list = self.df_orders['order_status'].unique()
        status_dictionary = {
            'Completed': 0,
            'Canceled': 0,
            'No show': 0,
            'Will call': 0,
        }

        checker(status_list, status_dictionary, 'order_status', 'Completed', 'Canceled', 'No show', 'Will call')

        completed = status_dictionary['Completed']
        canceled = status_dictionary['Canceled']
        no_show = status_dictionary['No show']
        will_call = status_dictionary['Will call']
        total_status_list = [completed, canceled, no_show, will_call]
        total_status = sum(total_status_list)

        self.status_table = pd.DataFrame({
            'Trip Status': ['Completed', 'Canceled', 'No Show', 'Will Call', 'Total'],
            'Count': [completed, canceled, no_show, will_call, total_status]
        })

        #print(self.status_table)
        

        #self.pivot_status = pd.pivot_table(self.df_orders, index=["order_status"], values=["order_id"], aggfunc="count").sort_values(by=["order_id"], ascending=False)
        #self.pivot_status.reset_index(inplace=True)
        #self.pivot_status = self.pivot_status.rename(columns= {"order_status": "Trip Status", "order_id": "Count"})
        
        #-----------------------OUTPUT---------------------------------

        #Cleaned TABLES
        self.driver_percentages = self.orders_by_driver_sum_sorted
        self.mode_table = self.mode_table
        self.trip_status_table = self.status_table

        
        #Formatted Tables
        self.driver_percentages.rename(columns={"full_name": 'Full Name', "orders_completed": 'Orders', "early_%(per_payer_settings)": 'Early %', "late_%(per_payer_settings)": 'Late %', "order_price": 'Revenue'}, inplace=True)

        orders_sum = sum(self.driver_percentages['Orders'])
        early_avg = round(self.driver_percentages['Early %'].mean(), 2)
        late_avg = round(self.driver_percentages['Late %'].mean(), 2)
        revenue_sum = self.driver_percentages['Revenue'].sum()


        canceled_sum = 0
        will_call_sum = 0

        for i in range(1, len(self.df_orders)):
            if self.df_orders.iloc[i]['order_status'] == 'Canceled':
                canceled_sum += self.df_orders.iloc[i]['order_price']
            elif self.df_orders.iloc[i]['order_status'] == 'Will Call':
                will_call_sum += self.df_orders.iloc[i]['order_price']

        


        self.driver_percentages.loc['Canceled'] = np.nan
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Full Name'] = 'Canceled'
        #self.driver_percentages.loc[self.driver_percentages.index[-1], ['Orders', 'Early %', 'Late %']] = ''
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Revenue'] = canceled_sum

        self.driver_percentages.loc['Will Call'] = np.nan
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Full Name'] = 'Will Call'
        #self.driver_percentages.loc[self.driver_percentages.index[-1], ['Orders', 'Early %', 'Late %']] = ''
        self.driver_percentages.loc[self.driver_percentages.index[-1], ['Revenue']] = will_call_sum


        self.driver_percentages.loc['Total'] = np.nan
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Full Name'] = 'Total'
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Orders'] = orders_sum
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Early %'] = early_avg
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Late %'] = late_avg
        self.driver_percentages.loc[self.driver_percentages.index[-1], 'Revenue'] = revenue_sum

        self.driver_percentages['Revenue'] = self.driver_percentages['Revenue'].apply(lambda x: '${0:,.2f}'.format(x))
        self.driver_percentages['Early %'] = self.driver_percentages['Early %'].apply(lambda x: format(x, '.2f'))
        self.driver_percentages['Late %'] = self.driver_percentages['Late %'].apply(lambda x: format(x, '.2f'))

        #removing NaNs
        self.driver_percentages.loc[self.driver_percentages.index[[-2, -3]], ['Early %', 'Late %']] = ''

        #print(self.driver_percentages)

        #RETURN OBJECT
        self.data = ReturnObject(
            self.date_range, 
            self.driver_percentages, 
            self.mode_table, 
            self.trip_status_table
        )

    def cleanUp(self,df):
        df.rename(columns=lambda x:x.replace("'", ''), inplace=True)
        df.rename(columns=lambda x:x.replace(" ", '_'), inplace=True)
        df.rename(columns=lambda x:x.lower(), inplace=True)

    def reformat(self, df):
        df.rename(columns=lambda x:x.replace('_', ' '))
        df.rename(columns=lambda x:x.title())

    
# Delete after testing
#class UploadFiles:
#    def __init__(self, x, y) -> None:
#        self.drivers = x
#        self.orders = y
#
#driver_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/drivers/driver_report.csv'
#order_file = 'C:/Users/Dispatch2/Desktop/DataAutomation/reports/daily/orders/order_report.csv'
#
#file_object = UploadFiles(driver_file, order_file)
#object = Analyze(file_object)


# Driver's Percentage 
#columns: first name, last name, orders completed, early%, late%

#Monthly Sales Report 
#columns: order ID, date, payer name, (driver) full name, order status, order price, order mode

#Cliff's Report
#columns: (driver) full name, order status, custom mode name, final price

#Where is safety score?