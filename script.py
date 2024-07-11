import openpyxl
import os
import shutil
import re
from fuzzywuzzy import fuzz
from PIL import Image

file_to_hanes_mapping = {
    "style_1": {
        "school_1": ["xsmall", "small", "medium", "large"],
        "school_2": ["xsmall", "small", "medium", "large"],
        "school_3": ["xsmall", "small", "medium", "large"],
        "school_4": ["xsmall", "small", "medium", "large"],
        "school_5": ["xsmall", "small", "medium", "large"],
        "school_n": ["xsmall", "small", "medium", "large"],
    },
    "style_2": {
        "school_1": ["xsmall", "small", "medium", "large"],
        "school_2": ["xsmall", "small", "medium", "large"],
        "school_3": ["xsmall", "small", "medium", "large"],
        "school_4": ["xsmall", "small", "medium", "large"],
        "school_5": ["xsmall", "small", "medium", "large"],
        "school_n": ["xsmall", "small", "medium", "large"],
    },
    "style_3": {
        "school_1": ["xsmall", "small", "medium", "large"],
        "school_2": ["xsmall", "small", "medium", "large"],
        "school_3": ["xsmall", "small", "medium", "large"],
        "school_4": ["xsmall", "small", "medium", "large"],
        "school_5": ["xsmall", "small", "medium", "large"],
        "school_n": ["xsmall", "small", "medium", "large"],
    },
    "style_4": {
        "school_1": ["xsmall", "small", "medium", "large"],
        "school_2": ["xsmall", "small", "medium", "large"],
        "school_3": ["xsmall", "small", "medium", "large"],
        "school_4": ["xsmall", "small", "medium", "large"],
        "school_5": ["xsmall", "small", "medium", "large"],
        "school_n": ["xsmall", "small", "medium", "large"],
    },    
    "style_5": {
        "school_1": ["xsmall", "small", "medium", "large"],
        "school_2": ["xsmall", "small", "medium", "large"],
        "school_3": ["xsmall", "small", "medium", "large"],
        "school_4": ["xsmall", "small", "medium", "large"],
        "school_5": ["xsmall", "small", "medium", "large"],
        "school_n": ["xsmall", "small", "medium", "large"],
    },
    "style_n": {
        "school_1": ["xsmall", "small", "medium", "large"],
        "school_2": ["xsmall", "small", "medium", "large"],
        "school_3": ["xsmall", "small", "medium", "large"],
        "school_4": ["xsmall", "small", "medium", "large"],
        "school_5": ["xsmall", "small", "medium", "large"],
        "school_n": ["xsmall", "small", "medium", "large"],
    }
}

table_data = [["school_1", "style1", "xsmall"],
              ["school_1", "style1", "small"],
              ["school_1", "style1", "medium"],
              ["school_1", "style1", "large"],

              ["school_2", "style1", "xsmall"],
              ["school_2", "style1", "small"],
              ["school_2", "style1", "medium"],
              ["school_2", "style1", "large"],
              
              ["school_3", "style1", "xsmall"],
              ["school_3", "style1", "small"],
              ["school_3", "style1", "medium"],
              ["school_3", "style1", "large"],

              ["school_4", "style1", "xsmall"],
              ["school_4", "style1", "small"],
              ["school_4", "style1", "medium"],
              ["school_4", "style1", "large"],


              ["school_1", "style2", "xsmall"],
              ["school_1", "style2", "small"],
              ["school_1", "style2", "medium"],
              ["school_1", "style2", "large"],

              ["school_2", "style2", "xsmall"],
              ["school_2", "style2", "small"],
              ["school_2", "style2", "medium"],
              ["school_2", "style2", "large"],
              
              ["school_3", "style2", "xsmall"],
              ["school_3", "style2", "small"],
              ["school_3", "style2", "medium"],
              ["school_3", "style2", "large"],

              ["school_4", "style2", "xsmall"],
              ["school_4", "style2", "small"],
              ["school_4", "style2", "medium"],
              ["school_4", "style2", "large"]]



def get_excel_rows(excel_file, sheet=None):
    workbook = openpyxl.load_workbook(excel_file)

    rows = []

    if not sheet:
        sheet = workbook.active
    else:
        sheet = workbook[sheet]

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            print(f"Warning: Skipping row with None value in school column: {row}")
            continue

        row_list = [row[0], row[1], str(row[2])]
        rows.append(row_list)
    
    return rows

def create_syle_map(excel_rows):
    style_map = {}

    for row in excel_rows:
        if row[1] not in style_map:
            style_map[row[1]] = {"schools": set()}

        if row[0] not in style_map[row[1]]:
            style_map[row[1]][row[0]] = []
        
        style_map[row[1]]["schools"].add(row[0])
        style_map[row[1]][row[0]].append(row[2])
    
    return style_map

def get_school_key_from_file(unique_schools, file_name):
    max_simliarity = -1
    answer = None

    for school in unique_schools:

        normalized_school = re.sub(r'[^a-z0-9]', ' ', school.lower()).strip()
        normalized_file_name = re.sub(r'[^a-z0-9]', ' ', file_name.lower()).strip()

        ratio = fuzz.ratio(normalized_school, normalized_file_name)

        if ratio > max_simliarity:
            max_simliarity = ratio
            answer = school

    return answer



def compress_image(image_path, compression_factor=75):
    picture = Image.open(image_path)
    picture.save(image_path, optimize=True, quality=compression_factor)


def copy_images(images_dir, styles_dir, style_map):
    
    for style in style_map:
        current_style_dir = os.path.join(styles_dir, style)
        current_image_dir = os.path.join(images_dir, style)
        image_files = os.listdir(current_image_dir)
        
        for image_file in image_files:
            file_school_name = get_school_key_from_file(style_map[style]["schools"], image_file)
            source_image_path = os.path.join(current_image_dir, image_file)
            
            for size_dir in style_map[style][file_school_name]:

                destination_image_path = os.path.join(current_style_dir, size_dir, image_file)
                os.makedirs(os.path.dirname(destination_image_path), exist_ok=True)
                shutil.copy(source_image_path, destination_image_path)

                print(f"Copied {source_image_path} to {destination_image_path}")
                
                compress_image(destination_image_path)


if __name__ == "__main__":
    rows = get_excel_rows("hanes_data.xlsx")
    style_map = create_syle_map(rows)

    copy_images("images", "styles", style_map)