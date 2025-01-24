import pandas as pd
from opencage.geocoder import OpenCageGeocode

def get_location(lat, lon, api_key, cache):
    # Check if the lat/lon pair is already cached
    if (lat, lon) in cache:
        return cache[(lat, lon)]
    
    # If not cached, make the API request
    geocoder = OpenCageGeocode(api_key)
    result = geocoder.reverse_geocode(lat, lon)
    if result:
        address = result[0]['formatted']
        cache[(lat, lon)] = address  # Save the result in cache
        return address
    else:
        cache[(lat, lon)] = "No address found"
        return "No address found"

# OpenCage API Key
api_key = "45aabb3771d24d208aa3c9a9ed4fb5d9"

# Load the Excel file
input_file = "input.xlsx"  # Replace with your file name
output_file = "output_with_addresses.xlsx"
sheet_name = "example"  # Replace with your sheet name if different

# Read the Excel file
df = pd.read_excel(input_file, sheet_name=sheet_name)

# Ensure latitude and longitude are numeric
df['Latitude'] = pd.to_numeric(df['LATITUDE'], errors='coerce')  # Column B
df['Longitude'] = pd.to_numeric(df['LONGITUDE'], errors='coerce')  # Column C

# Drop rows with invalid latitude/longitude
df = df.dropna(subset=['Latitude', 'Longitude'])

# Create a cache for API results
cache = {}

# Get addresses and add them to column J
df['Address'] = df.apply(
    lambda row: get_location(row['Latitude'], row['Longitude'], api_key, cache), axis=1
)

# Place the "Address" column in column J (assuming original columns go up to column I)
columns = list(df.columns)
if 'Address' in columns:
    columns.remove('Address')
columns.insert(9, 'Address')  # Column J is the 10th column (index 9)

# Reorder and save the results to a new Excel file
df = df[columns]
df.to_excel(output_file, index=False)

print(f"Results saved to {output_file}")
