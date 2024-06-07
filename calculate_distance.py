import pandas as pd
import googlemaps
import time

# Configure your Google Maps API
API_KEY = 'GOOGLE MAPS API KEY'
gmaps = googlemaps.Client(key=API_KEY)
measure_units = 'imperial' # You can choose between imperial or metric

# Configure your excel file parameters
source_excel_file = "TYPE YOUR SOURCE FILE PATH/NAME" # Excel file where address data are
target_excel_file = "TYPE YOUR TARGET FILE PATH/NAME" # Excel file to save output with distance measures
source_column_name = 'Source' # Column name inside excel file that has source address
target_column_name = 'Target' # Column name inside excel file that has target address
distance_column_name = 'Distance' # Column name that will be created inside excel file to register distance

# Load the Excel Spreadsheet in a pandas dataframe
df = pd.read_excel(source_excel_file)

# Add the column to register the distance
df[distance_column_name] = None

# Function to calculate the distance using Google Maps Distance API
def calculate_distance(source, destination):
    try:
        result = gmaps.distance_matrix(source, destination, mode='driving', units=measure_units)
        distance = result['rows'][0]['elements'][0]['distance']['text']
        return distance
    except Exception as e:
        print(f"Error calculating the distance between {source} and {destination}: {e}")
        return None

# Loop to iterate over all lines and calculate the distance for each address pair
for index, row in df.iterrows():
    source = row['Target']
    destination = row['Source']
    df.at[index, 'Distance'] = calculate_distance(source, destination)
    time.sleep(0.05)  # Wait 50 ms to avoid API quota problem

# Save Dataframe in a new excel file
df.to_excel('distances_calculadas.xlsx', index=False)