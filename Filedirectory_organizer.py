import os
import shutil

def organize_files(input_directory, output_directory):
    # Ensure the output directory exists
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Iterate over all files in the input directory
    for filename in os.listdir(input_directory):
        if os.path.isfile(os.path.join(input_directory, filename)):
            # Remove the extension from the filename
            filename_without_extension = os.path.splitext(filename)[0]
            
            # Split the filename based on the "|" character
            parts = filename_without_extension.split('|')
            
            if len(parts) == 2:
                afdeling = parts[0].strip()
                eigenaar = parts[1].strip()

                # Create Verzender directory structure
                verzender_dir = os.path.join(output_directory, "Verzender", afdeling)
                if not os.path.exists(verzender_dir):
                    os.makedirs(verzender_dir)

                # Copy the file to the Verzender directory and rename it
                verzender_filename = f"Verzenden aan {eigenaar}.xlsx"
                shutil.copy(os.path.join(input_directory, filename), os.path.join(verzender_dir, verzender_filename))

                # Create Ontvanger directory structure
                ontvanger_dir = os.path.join(output_directory, "Ontvanger", eigenaar)
                if not os.path.exists(ontvanger_dir):
                    os.makedirs(ontvanger_dir)

                # Copy the file to the Ontvanger directory and rename it
                ontvanger_filename = f"Ontvangen van {afdeling}.xlsx"
                shutil.copy(os.path.join(input_directory, filename), os.path.join(ontvanger_dir, ontvanger_filename))