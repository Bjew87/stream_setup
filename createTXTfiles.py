import openpyxl
import os


import json


import sys


def read_xlsx(base_path, input_file, target_file):

    # reading file

    file_path = str(input_file)

    wb = openpyxl.load_workbook(file_path)

    ws = wb.active

    # load the template scene data

    with open(base_path+'\\json_templates\\script_template_scenes.json') as json_file:

        data = json.load(json_file)
        head, tail = os.path.split(target_file)
        data['name'] = tail.replace('.json', '')
    #
    for idx, row in enumerate(ws.iter_rows()):

        if idx > 0:

            #

            scene_data = create_scene_data(base_path, ws, idx)

            #

            text_data = create_text_data(base_path, ws, idx)

            #

            data['sources'].append(scene_data)
            data['sources'].append(text_data)

    # write final data to target file

    with open(target_file, 'w') as f:

        json.dump(data, f, indent=4)


def create_scene_data(base_path, ws, idx):

    with open(base_path+'\\json_templates\\scene_template.json') as json_file:

        data = json.load(json_file)

        data['name'] = ws.cell(idx+1, 1).value.replace(' ', '_')

        data['settings']['items'][3]['name'] = ws.cell(

            idx+1, 1).value.replace(' ', '_') + "_text"

        return data


def create_text_data(base_path, ws, idx):

    with open(base_path+'\\json_templates\\text_template.json') as json_file:

        data = json.load(json_file)

        data['name'] = ws.cell(idx+1, 1).value.replace(' ', '_') + "_text"

        data['settings']["text"] = ws.cell(idx+1, 2).value
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
