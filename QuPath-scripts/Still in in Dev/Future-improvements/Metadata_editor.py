# Requirement for script: Exif. run pip install exif

from exif import Image
import os

folder_path  = r"C:\Users\Bob\OneDrive\Desktop\Third QuPath\copied_images"


img_filename = []

img_path= []

for i in os.listdir(folder_path):
    if i.endswith("tif"):
        img_filename.append(i)


for i in os.listdir(folder_path):
    if i.endswith("tif"):
        img_filename.append(i)
        full_path = os.path.join(folder_path, i)
        img_path.append(full_path)


for i in img_path:
    try:
        with open(i, 'rb') as img_file:
            img = Image(img_file)
            print(img.has_exif)
            sorted(img.list_all())
    except Exception as e:
        print(f'{i} does not have metadata. error: {e}')
