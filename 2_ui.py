# Import necessary libraries
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import datetime, time, timedelta
import matplotlib.pyplot as plt
#import os
import streamlit as st
import io
import base64
#####
########           ADDED           ####################

def int_to_ampm(hour_int):
    # Convert the integer hour to a string with AM/PM
    from datetime import datetime
    return datetime.strptime(str(hour_int), "%H").strftime("%I%p")

def handle_multiday_meetings(df):
    new_rows = []
    indices_to_remove = []

    for index, row in df.iterrows():
        start_datetime = row['start_datetime']
        end_datetime = row['end_datetime']
        if start_datetime.date() == end_datetime.date():
            continue
        else:
            indices_to_remove.append(index)
            
            # For the first day
            first_row = row.copy()
            first_row['end_datetime'] = start_datetime.replace(hour=23, minute=59, second=59)
            new_rows.append(first_row)
            print(f"Added date: {start_datetime.date()}")

            # For the subsequent days
            current_date = start_datetime + pd.Timedelta(days=1)
            while current_date.date() < end_datetime.date():
                new_row = row.copy()
                new_row['start_datetime'] = current_date.replace(hour=0, minute=0, second=0)
                new_row['end_datetime'] = current_date.replace(hour=23, minute=59, second=59)
                new_rows.append(new_row)
                print(f"Added date: {current_date.date()}")
                current_date += pd.Timedelta(days=1)

            # For the last day
            last_row = row.copy()
            last_row['start_datetime'] = end_datetime.replace(hour=0, minute=0, second=0)
            new_rows.append(last_row)
            print(f"Added date: {end_datetime.date()}")

    # Remove the original multi-day meeting rows
    df.drop(indices_to_remove, inplace=True)
    
    # Append the new adjusted rows
    df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    return df


# Function to calculate difference considering only times between 8AM-6PM
def calculate_difference(row):
    start = row['start_datetime']
    print("Start datetime: ", start)
    end = row['end_datetime']
    print("End datetime: ", end)
    
    # Define working hours
    print("Start hour: ", start_hour)
    print("End hour: ", end_hour)
    work_start_time = time(start_hour, 0)
    print("Work start hour: ", work_start_time)
    work_end_time = time(end_hour, 0)
    print("Work end hour: ", work_end_time)
    
    # Adjust start and end times if outside of working hours
    if start.time() < work_start_time:
        start = datetime.combine(start.date(), work_start_time) #make it the start if start is actually earlier
        print("New start time is: ", start)
    if end.time() > work_end_time:
        end = datetime.combine(end.date(), work_end_time)  #make it the end if end is actually later
        print("New end time is: ", end)

    # Calculate total time within working hours across days
    total_seconds = 0
    current = start
    while current < end:
        next_bound = datetime.combine(current.date(), work_end_time)
        if next_bound > end:
            next_bound = end
        if next_bound > current:  # Ensure it's within working hours
            total_seconds += (next_bound - current).total_seconds()
        current = datetime.combine(current.date() + timedelta(days=1), work_start_time)
    
    total_time_hours = total_seconds / 3600
    print("Total time within bounds: ", total_time_hours)

    if total_time_hours >= total_time_per_day:
        total_time_hours = total_time_per_day
        
    return total_time_hours  # Convert seconds to hours

########           ADDED           ####################
# Title
st.title("Space Utilization Report Generator")


# Date Range Input
start_date = st.date_input("Start Date", datetime.today())
end_date = st.date_input("End Date", datetime.today())

# Hour Range Input
start_hour = st.selectbox("Start Hour", list(range(0, 24)))
end_hour = st.selectbox("End Hour", list(range(0, 24)))

#define total time
total_time_per_day = end_hour-start_hour

start_hour_string = int_to_ampm(start_hour)
end_hour_string = int_to_ampm(end_hour)


# Days of Week Checkbox
days_of_week = [
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"
]

selected_days = []
for day in days_of_week:
    if st.checkbox(day):
        selected_days.append(day)
# Upload a file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "csv"])
# Check if a file has been uploaded
if uploaded_file is not None:
    # Check if the uploaded file is not empty
    if uploaded_file.size > 0:
        # Read in the DataFrame
        df = pd.read_csv(uploaded_file)
        # You can now work with the DataFrame 'df'
        st.write(df)
        df_meetings = df[['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'Meeting Organizer']]
        #Convert the necessary data types
        df_meetings['Start Date'] = pd.to_datetime(df_meetings['Start Date'])
        df_meetings['End Date'] = pd.to_datetime(df_meetings['End Date'])
        df_meetings['Start Time'] = df_meetings['Start Time'].astype(str)
        df_meetings['End Time'] = df_meetings['End Time'].astype(str)

        # Convert dates to string and then combine them with times to form datetime objects
        df_meetings.loc[:, 'start_datetime'] = pd.to_datetime(df_meetings['Start Date'].dt.strftime('%Y-%m-%d') + ' ' + df_meetings['Start Time'].astype(str))
        df_meetings.loc[:, 'end_datetime'] = pd.to_datetime(df_meetings['End Date'].dt.strftime('%Y-%m-%d') + ' ' + df_meetings['End Time'].astype(str))

        # Assuming your df_meetings dataframe is defined elsewhere in your code.
        df_meetings = handle_multiday_meetings(df_meetings)
        # Calculate total_time
        df_meetings['total_time'] = df_meetings.apply(calculate_difference, axis=1)
        average_meeting_length = df_meetings["total_time"].mean()
        date_range = pd.date_range(start=start_date, end=end_date) #this should be min and max of input data
        utilization_df = pd.DataFrame({'date': date_range})

        # 2. Calculate utilization for each date
        utilization = df_meetings.groupby('Start Date')['total_time'].sum()

        # Cap the utilization at 10 after grouping
        utilization = utilization.apply(lambda x: total_time_per_day if x > total_time_per_day else x)

        # 3. Merge the results into the final_df
        utilization_df = utilization_df.merge(utilization, left_on='date', right_index=True, how='left')
        utilization_df.fillna(0, inplace=True)  # Fill NA values with 0 for dates with no utilization
        # Remove weekends
        utilization_df = utilization_df[~utilization_df['date'].dt.weekday.isin([5, 6])]

        # List of specific dates to exclude - maybe do all calendar dates over the past 3 years
        exclude_dates = [
            "2022-05-30",
            "2022-06-20",
            "2022-07-04",
            "2022-09-05",
            "2022-11-24",
            "2022-11-25",
            "2022-12-26",
            "2022-12-27",
            "2022-12-28",
            "2022-12-29",
            "2022-12-30",
            "2023-01-02",
            "2023-01-16",
            "2023-02-20",
            "2023-05-29",
            "2023-06-19",
            "2023-07-03",
            "2023-07-04"
        ]

        # Remove specific dates
        utilization_df = utilization_df[~utilization_df['date'].isin(pd.to_datetime(exclude_dates))]

        df_meetings = df_meetings[~df_meetings['start_datetime'].isin(pd.to_datetime(exclude_dates))]
        num_hours_daily = end_hour-start_hour

        utilization_df['percent_utilized'] = (utilization_df['total_time'] / num_hours_daily) * 100

        utilization_df['weekday'] = utilization_df['date'].dt.day_name()

        # Filter only weekdays (Monday to Friday)
        weekday_df = utilization_df[utilization_df['weekday'].isin(selected_days)]

        # Convert 'weekday' column to categorical type with ordered categories
        ordered_days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
        weekday_df['weekday'] = pd.Categorical(weekday_df['weekday'], categories=ordered_days, ordered=True)

        # Aggregate by weekday
        weekday_aggregation = weekday_df.groupby('weekday').agg(
            average_percent_utilized=('percent_utilized', 'mean')
        ).reset_index()

        # Aggregate by Month+Year
        utilization_df['month_year'] = utilization_df['date'].dt.to_period('M')
        month_aggregation = utilization_df.groupby('month_year').agg(
            average_percent_utilized=('percent_utilized', 'mean')
        ).reset_index()

    else:
        st.warning("Uploaded file is empty.")
else:
    st.info("Please upload an Excel file.")


# Assuming the DataFrames have already been created as utilization_df, month_aggregation, and weekday_aggregation

    # Function to download data as csv or excel
def download_link(object_to_download, download_filename, download_link_text):
    if isinstance(object_to_download, pd.DataFrame):
        object_to_download = object_to_download.to_csv(index=False)

    # some strings <-> bytes conversions necessary here
    b64 = base64.b64encode(object_to_download.encode()).decode()
    return f'<a href="data:file/txt;base64,{b64}" download="{download_filename}">{download_link_text}</a>'

# Display the DataFrame
#st.write(df)

# Create download link for Excel with multiple sheets
if st.button('Download Excel'):
    towrite = io.BytesIO()
    
    # Create an Excel writer object
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Raw Data')
        df_meetings.to_excel(writer, index=False, sheet_name='Meetings Data')
        utilization_df.to_excel(writer, index=False, sheet_name='Utilization')
        month_aggregation.to_excel(writer, sheet_name='Month Aggregation', index=False)
        weekday_aggregation.to_excel(writer, sheet_name='Weekday Aggregation', index=False)

    
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode() 
    link=f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="report">Download excel file</a>'
    st.markdown(link, unsafe_allow_html=True)
