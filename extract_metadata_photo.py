import os
from PIL import Image
import exifread
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
import openpyxl

def get_decimal_from_dms(dms, ref):
    degrees = dms[0].num / dms[0].den
    minutes = dms[1].num / dms[1].den
    seconds = dms[2].num / dms[2].den

    decimal = degrees + (minutes / 60.0) + (seconds / 3600.0)
    if ref in ['S', 'W']:
        decimal = -decimal
    return decimal

def reverse_geocode(lat, lon):
    try:
        # Use geopy to get location details
        geolocator = Nominatim(user_agent="photo_metadata_extractor")
        location = geolocator.reverse((lat, lon), exactly_one=True)
        return location
    except (GeocoderTimedOut, GeocoderServiceError) as e:
        print(f"Geocoding error: {e}")
        return None

def extract_metadata(image_path):
    metadata = {
        'filename': os.path.basename(image_path),
        'country': 'N/A',
        'state': 'N/A',
        'city': 'N/A',
        'datetime': 'N/A',
        'device_id': 'N/A',
        'imei': 'N/A'
    }

    # Open image using Pillow
    image = Image.open(image_path)

    # Open image file for reading (binary mode)
    with open(image_path, 'rb') as img_file:
        # Return Exif tags
        tags = exifread.process_file(img_file)

    # Extract GPS coordinates
    gps_latitude = tags.get('GPS GPSLatitude')
    gps_latitude_ref = tags.get('GPS GPSLatitudeRef')
    gps_longitude = tags.get('GPS GPSLongitude')
    gps_longitude_ref = tags.get('GPS GPSLongitudeRef')

    if gps_latitude and gps_latitude_ref and gps_longitude and gps_longitude_ref:
        lat = get_decimal_from_dms(gps_latitude.values, gps_latitude_ref.values)
        lon = get_decimal_from_dms(gps_longitude.values, gps_longitude_ref.values)

        location = reverse_geocode(lat, lon)
        if location:
            address = location.raw.get('address', {})
            metadata['country'] = address.get('country', 'N/A')
            metadata['state'] = address.get('state', 'N/A')
            metadata['city'] = address.get('city', 'N/A')

    # Extract DateTime
    datetime = tags.get('EXIF DateTimeOriginal')
    if datetime:
        metadata['datetime'] = datetime.values

    # Extract Device ID and IMEI if available
    device_id = tags.get('Image HostComputer')
    if device_id:
        metadata['device_id'] = device_id.values

    imei = tags.get('EXIF IMEINumber')  # This might vary based on the camera, adjust accordingly
    if imei:
        metadata['imei'] = imei.values

    return metadata

def process_photos(folder_path):
    photos_metadata = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('jpg', 'jpeg', 'png')):
            file_path = os.path.join(folder_path, filename)
            metadata = extract_metadata(file_path)
            photos_metadata.append(metadata)

    return photos_metadata

def create_excel(metadata_list, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Photos Metadata"

    headers = ['Filename', 'Country', 'State', 'City', 'Date and Time', 'Device ID', 'IMEI']
    sheet.append(headers)

    for metadata in metadata_list:
        row = [
            metadata['filename'],
            metadata['country'],
            metadata['state'],
            metadata['city'],
            metadata['datetime'],
            metadata['device_id'],
            metadata['imei']
        ]
        sheet.append(row)

    workbook.save(output_file)

# Main script
folder_path = 'photos'  # Change to your folder containing photos
output_file = 'photos_metadata.xlsx'
metadata_list = process_photos(folder_path)
create_excel(metadata_list, output_file)

print(f"Metadata extracted and saved to {output_file}")
