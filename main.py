import pandas as pd

# Read the Excel file
data = pd.read_excel('Assignment_Timecard.xlsx')

# Drop unused columns and duplicates
data = data.drop(['Unnamed: 9', 'Unnamed: 10'], axis=1)
data = data.drop_duplicates()

# Get dates from datetime
data['Time_in_date'] = pd.to_datetime(data['Time']).dt.date
data['Time_out_date'] = pd.to_datetime(data['Time Out']).dt.date

# Get next and previous timein by an employee
data.sort_values(['Employee Name', 'Time'], inplace=True)
data['next_timein'] = data.groupby('Employee Name')['Time'].shift(-1)
data['prev_timein'] = data.groupby('Employee Name')['Time'].shift(1)
data['prev_timein_date'] = data.groupby('Employee Name')['Time_in_date'].shift(1)

# Finding employees worked on 7 consecutive days
task1_data = data[data['Time_in_date'] != data['prev_timein_date']]
task1_data['days_diff'] = (pd.to_datetime(task1_data['Time_in_date']) - pd.to_datetime(task1_data['prev_timein_date'])).dt.days
task1_data['consecutive_group'] = (task1_data['days_diff'] != 1.0).cumsum()

task1_usrs = task1_data.groupby(['Employee Name', 'consecutive_group']).filter(lambda x: len(x) >= 7)
task1_result = task1_usrs[['Employee Name', 'Position ID']].drop_duplicates()

# Finding employees who have less than 10 hours of time between shifts but greater than 1 hour
data['time_btw_shifts'] = (data['next_timein'] - data['Time']).dt.total_seconds()/3600
task2_data = data[(data['time_btw_shifts'] < 10.0) & (data['time_btw_shifts'] > 1.0)]
task2_result = task2_data[['Employee Name', 'Position ID']].drop_duplicates()

# Finding employees who has worked for more than 14 hours in a single shift
data['time_shift'] = (data['Time Out'] - data['Time']).dt.total_seconds()/3600
task3_data = data[data['time_shift'] > 14.0]
task3_result = task3_data[['Employee Name', 'Position ID']].drop_duplicates()

# Writing output to text file
task3_result.to_csv('output3.txt', sep='\t', index=False)
task1_result.to_csv('output1.txt', sep='\t', index=False)
task2_result.to_csv('output2.txt', sep='\t', index=False)
