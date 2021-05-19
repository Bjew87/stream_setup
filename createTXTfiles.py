#!utf8

"""

        ########        ## ######## ##      ## 
        ##     ##       ## ##       ##  ##  ## 
        ##     ##       ## ##       ##  ##  ## 
        ########        ## ######   ##  ##  ## 
        ##     ## ##    ## ##       ##  ##  ## 
        ##     ## ##    ## ##       ##  ##  ## 
        ########   ######  ########  ###  ###  


    - Script reads Excel file with scene title and text to display and creates OBS scenes.

    - 

"""
import openpyxl
import os
import json
import sys

from openpyxl.descriptors.base import Length

def read_xlsx(base_path, input_file, target_file):
    # reading excel file
    file_path = str(input_file)
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    # load the template scene collection data
    with open(base_path+'\\json_templates\\scene_collection_template.json', encoding='utf8') as json_file:
        data = json.load(json_file)
        head, tail = os.path.split(target_file)
        data['name'] = tail.replace('.json', '')
    # set default variables
    big_text_length = 175
    # read all rows from excel file
    for idx, row in enumerate(ws.iter_rows()):
        # iterate through all row in the excel except header row
        if idx > 0:
            # if there is data written
            if ws.cell(idx+1, 1).value:
                # start creation of scenes
                scene_title = "[S]__"+ ws.cell(idx+1, 1).value.replace(' ', '_')
                #
                print('Creating Scene: ' + scene_title)
                #
                scene_text_header_ID = scene_title + "_text_header"
                scene_text_ID = scene_title + "_text"
                scene_text_header = ws.cell(idx + 1, 2).value
                scene_text = ws.cell(idx + 1, 3).value
                font_size_line_1 = 64
                font_size_line_2 = 92
                posY = 1145
                text_length = len(scene_text)
                color = None
                #
                if(text_length >= big_text_length):
                    font_size_line_1 = 72
                    font_size_line_2 = 64
                    color = 4278190080
                # only text has been given
                if scene_text is None or (scene_text_header is None and scene_text is not None):
                    font_size_line_1 = font_size_line_2
                    posY = 1170
                # only second column has been filled
                # write text to first header and set text line to None
                if scene_text_header is None and scene_text is not None:
                    scene_text_header = scene_text
                    scene_text = None
                #
                text_header_data = create_text_data(
                    base_path, scene_text_header_ID, scene_text_header, font_size_line_1, color)
                #
                if scene_text is not None:
                    text_data = create_text_data(
                        base_path, scene_text_ID, scene_text, font_size_line_2, color)
                #
                if(text_length < big_text_length):
                    scene_data = create_scene_data(
                        base_path, scene_title, scene_text_ID, scene_text_header_ID, posY)
                else:
                    #
                    scene_data = create_scene_data_big_text(
                        base_path, scene_title, scene_text_ID, scene_text_header_ID)
                # add all created JSON block to the scene collection
                data['sources'].append(scene_data)
                data['sources'].append(text_header_data)
                data['sources'].append(text_data)
    # we gathered all data and appended it to the template scene collection
    with open(target_file, 'w') as f:
        #
        print('Writing scene collection to: ' + target_file)
        # write final scene collection to target file
        json.dump(data, f, indent=4)


def create_scene_data_big_text(base_path, scene_title, scene_text_ID, scene_text_header_ID):
    with open(base_path+'\\json_templates\\scene_big_text_template.json', encoding='utf8') as json_file:
        data = json.load(json_file)
        data['name'] = scene_title
        data['settings']['items'][7]['name'] = scene_text_ID
        data['settings']['items'][8]['name'] = scene_text_header_ID
        return data

def create_scene_data(base_path, scene_title, scene_text_ID, scene_text_header_ID, posX):
    with open(base_path+'\\json_templates\\scene_template.json', encoding='utf8') as json_file:
        data = json.load(json_file)
        data['name'] = scene_title
        data['settings']['items'][4]['name'] = scene_text_ID
        data['settings']['items'][5]['name'] = scene_text_header_ID
        data['settings']['items'][5]['pos']['y'] = posX
        return data


def create_text_data(base_path, scene_text_ID, scene_text, font_size, color):
    with open(base_path+'\\json_templates\\text_template.json', encoding='utf8') as json_file:
        data = json.load(json_file)
        data['name'] = scene_text_ID
        data['settings']["text"] = scene_text
        data['settings']["font"]['size'] = font_size
        if color:
            data['settings']['color'] = color
        return data


if __name__ == "__main__":
    #
    print('Welcome to the automated scene creation by BJEW')
    #
    if sys.argv and len(sys.argv) > 2:
        #
        input_file = sys.argv[1]
        target_file = sys.argv[2]
        #
        read_xlsx(os.path.dirname(sys.argv[0]), input_file, target_file)
        #
        print('Success!')
    else:
        #
        print('Missing parameters!')
