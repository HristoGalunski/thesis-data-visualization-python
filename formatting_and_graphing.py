import pandas as pd
import matplotlib.pyplot as plt



#Extracting the values inside the excel file
PATH = "C:\\Users\\Ico\\Desktop\\Dev\\Python\\Thesis\\sensor_data.xlsx"
excel_df = pd.read_excel(PATH)


#Isolating only the specific columns for each sensor and joining them into one dataframe
#Sensor0
data0 = excel_df['Data s0'].to_frame()
time0 = excel_df['Time s0'].to_frame()
df0 = time0.join(data0).dropna(subset = ['Data s0'])

#Sensor1
data1 = excel_df['Data s1'].to_frame()
time1 = excel_df['Time s1'].to_frame()
df1 = time1.join(data1).dropna(subset = ['Data s1'])

#Sensor2
data2 = excel_df['Data s2'].to_frame()
time2 = excel_df['Time s2'].to_frame()
df2 = time2.join(data2).dropna(subset = ['Data s2'])


#Building a dictionary to use as a way to translate the month values from the file's metadata to numbers
month = {'Jan' : '1',
                 'Feb' : '2',
                 'Mar' : '3',
                 'Apr' : '4',
                 'May' : '5',
                 'Jun' : '6',
                 'Jul' : '7',
                 'Aug' : '8',
                 'Sep' : '9',
                 'Oct' : '10',
                 'Nov' : '11',
                 'Dec' : '12',}


#Function to iterate through each row of the date columns the sensors to convert them to a
# format that we can use for making the graphs
def formatDates(dataframe, sensor_number):
    time_dataframe = dataframe['Time s' + sensor_number]
    
    for i in range(len(time_dataframe)):
        elem = time_dataframe[i].split()
        time_dataframe[i] = '{}-{}-{} {}'.format(elem[4], month.get(elem[1]), elem[2], elem[3])

    return dataframe


def plotDataFrame(dataframe, sensor_number):
    plt.plot_date(dataframe['Time s' + sensor_number],
                            dataframe['Data s' + sensor_number], linestyle='solid')
    
    plt.gcf().autofmt_xdate()
    plt.show()


#Formatting all the dataframes
df0 = formatDates(df0, '0')
df1 = formatDates(df1, '1')
df2 = formatDates(df2, '2')
           

plotDataframe(df0, '0')
plotDataframe(df1, '1')
plotDataframe(df2, '2')



    
