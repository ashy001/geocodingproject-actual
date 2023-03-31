import gmaps
import gmaps.datasets
import pandas as pd
import googlemaps

# Load the Excel data
data = pd.read_excel('file.xlsx')

# Set up the Google Maps API client
with open('api_key.txt', 'r') as file:
    api_key = file.read()
gmaps_client = googlemaps.Client(key=api_key)

# Geocode the addresses
geocoded_data = []
for index, row in data.iterrows():
    address = row['Address']
    geocode_result = gmaps_client.geocode(address)
    if geocode_result:
        lat = geocode_result[0]['geometry']['location']['lat']
        lng = geocode_result[0]['geometry']['location']['lng']
    else:
        lat = None
        lng = None
    geocoded_data.append({
        'Name': row['Name'],
        'Address': address,
        'Phone Number': row['Phone Number'],
        'Latitude': lat,
        'Longitude': lng
    })

# Create the map
geocoded_locations = [(d['Latitude'], d['Longitude']) for d in geocoded_data]
gmaps.configure(api_key=api_key)
fig = gmaps.Map(center=(geocoded_locations[0][0], geocoded_locations[0][1]), zoom_level=12, map_type='SATELLITE')
marker_locations = gmaps.marker_layer(geocoded_locations)
fig.add_layer(marker_locations)

# Apply custom styles to the map
custom_style = [{
    'featureType': 'poi',
    'elementType': 'labels',
    'stylers': [{'visibility': 'off'}]
}, {
    'featureType': 'poi',
    'elementType': 'geometry',
    'stylers': [{'visibility': 'off'}]
}]
custom_map = gmaps.Map(map_type='SATELLITE', zoom_level=12, center=(geocoded_locations[0][0], geocoded_locations[0][1]), styles=custom_style)
fig.styled_map = custom_map

# Save the geocoded data
geocoded_data_df = pd.DataFrame(geocoded_data)
geocoded_data_df.to_excel('geocoded_data.xlsx', index=False)
