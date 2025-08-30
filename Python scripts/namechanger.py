import os

directory = r"C:\Users\psoor\OneDrive\Desktop\Bachelorarbeit\Ductal-cells-QuPath\Ductal_measurements_all_ducts"

to_change = "40x"
change_t0 = "01.07"

for image in os.listdir(directory):
    if to_change in image:
        # Create full file paths
        old_filepath = os.path.join(directory, image)
        
        # Replace the text in the filename
        new_filename = image.replace(to_change, change_t0)
        new_filepath = os.path.join(directory, new_filename)
        
        # Rename the file
        try:
            os.rename(old_filepath, new_filepath)
            print(f"Renamed: {image} -> {new_filename}")
        except OSError as e:
            print(f"Error renaming {image}: {e}")


