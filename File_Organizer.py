import os
import time
import shutil
from PIL import Image, ExifTags

from PIL.ExifTags import TAGS
from pillow_heif import HeifImage
from hachoir.parser import createParser
from hachoir.metadata import extractMetadata

from PIL import UnidentifiedImageError
# Register HEIC support
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    print("⚠️ 'pillow-heif' not installed. HEIC support will not work without it.")

# Get EXIF date for images
def get_exif_date(path):
    try:
        image = Image.open(path)

        if hasattr(image, "_getexif"):
            exif_data = image._getexif()
            if exif_data:
                date_fields = ['DateTimeOriginal', 'DateTimeDigitized', 'DateTime']
                for tag_id, value in exif_data.items():
                    tag = TAGS.get(tag_id, tag_id)
                    if tag in date_fields:
                        try:
                            return time.strptime(value, '%Y:%m:%d %H:%M:%S')
                        except ValueError:
                            continue
    except UnidentifiedImageError:
        print(f"Unrecognized image: {path}")
        return None
    except Exception as e:
        print(f"Error reading EXIF from {path}: {e}")
        return None
    
    return None

# Get the video creation date
def get_mov_creation_date(path):
    parser = createParser(path)
    if not parser:
        print("Unable to parse file")
        return None

    with parser:
        metadata = extractMetadata(parser)
        if metadata:
            creation_date = metadata.get("creation_date")
            return creation_date
        else:
            print("No metadata found")
            return None


# Check that the date was valid
def validDate(date_taken, file):
        if not date_taken:
            print(f"Skipping {file} (no valid date found).")
            return False
        
        return True


def move_normal_photos(date_taken, file, full_path):
            
            if not validDate(date_taken, file):
                folder_name =  "Unkown Date"

            else:
                folder_name = f"{date_taken.tm_year}-{date_taken.tm_mon:02d}"
                
            destination_folder = os.path.join(folder_path, folder_name)
                
            os.makedirs(destination_folder, exist_ok=True)

            # Prevent overwriting
            new_path = os.path.join(destination_folder, file)
            if os.path.exists(new_path):
                base, ext = os.path.splitext(file)
                count = 1
                while os.path.exists(new_path):
                    new_filename = f"{base}_{count}{ext}"
                    new_path = os.path.join(destination_folder, new_filename)
                    count += 1

            # Attempts to move the file
            try:
                shutil.move(full_path, new_path)
                print(f"Moved: {file} -> {folder_name}")
            except Exception as e:
                print(f"Error moving {file}: {e}")


# Converts string to struct_time object
def convert_string_to_date(stringDate):

    arrayDate = stringDate.replace(" ", ":")
    x = arrayDate.split(":")
    date = time.struct_time((
        int(x[0]), # Year
        int(x[1]), # Month
        int(x[2]), # Day
        int(x[3]), # Hour
        int(x[4]), # Min
        int(x[5]), # Second
        0,
        0,
        -1
    ))
    return date

# Converts a date time object to struct_time object
def convert_datetime_to_date(dateTime):
    date = time.struct_time((
        int(dateTime.year), # Year
        int(dateTime.month), # Month
        int(dateTime.day), # Day
        int(dateTime.hour), # Hour
        int(dateTime.minute), # Min
        int(dateTime.second), # Second
        0,
        0,
        -1
    ))

    return date

# Sort the supported files in the directory
def sort_photos_by_date_taken(folder_path):
    if not os.path.exists(folder_path):
        print(f"Folder '{folder_path}' does not exist.")
        return

    # All the supported extensions
    supported_extensions = ('.jpg', '.jpeg', '.png', '.heic', '.HEIC', '.mov', '.MOV', '.mp4', '.3GP', '.3gp')

    # Loops through all the files in the given folder
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            
            if not file.lower().endswith(supported_extensions):
                continue
            
            full_path = os.path.join(root, file)

            # Optional: only sort files from the main folder
            if root != folder_path:
                continue

            # Checks that the files exists
            if not os.path.isfile(full_path):
                print(f"File not found or inaccessible: {full_path}")
                continue

            date_taken = None

            # Handle image files
            if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                date_taken = get_exif_date(full_path)
            
            # Handle video files
            elif file.lower().endswith(('.mov', '.MOV', '.mp4', '.3gp')):
                date_taken = get_mov_creation_date(full_path)

                if date_taken is None:
                    continue
                date_taken = convert_datetime_to_date(date_taken)

            # Handles Heic files
            elif file.lower().endswith(('.heic', '.HEIC')):
                img = Image.open(full_path)
                exif = img.getexif()
                # Find the tag ID for 'DateTime'
                for tag_id, value in exif.items():
                    tag = ExifTags.TAGS.get(tag_id, tag_id)
                    if tag == 'DateTime':
                        date_taken = convert_string_to_date(value)
                        break
                # Closes the file
                img.close()
            

            # Moves the the file to normal location
            move_normal_photos(date_taken, file, full_path)


if __name__ == "__main__":
    folder_path = input("Enter the path to the folder: ").strip()
    sort_photos_by_date_taken(folder_path)
