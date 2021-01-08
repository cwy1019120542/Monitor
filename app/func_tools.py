import re
import os
import uuid
import pandas
import zipfile
import json
import shutil
import requests
import aiohttp
import rarfile
from urllib3 import encode_multipart_formdata
from .response import parameter_error
from urllib.parse import quote
from flask import send_from_directory, make_response, Response
from .parameter_config import zip_suffix

def to_xlsx(file_path):
    file_name, suffix = os.path.splitext(file_path)
    result_file_name = f'{file_name}.xlsx'
    if suffix == ".XLSX":
        os.rename(file_path, result_file_name)
    if suffix not in (".xlsx", ".XLSX"):
        excel_file = pandas.read_excel(file_path)
        excel_file.to_excel(result_file_name, index=False)
        os.remove(file_path)
    return result_file_name

def clean_file_name(file_name):
    sub_list = r"[/\\\:\*\?\"\<\>\|-]"
    new_file_name = re.sub(sub_list, '_', str(file_name))
    return new_file_name

def parameter_check(request_data, parameter_group_list, is_all=True):
    if request_data == None:
        return False, parameter_error([i[0] for i in parameter_group_list])
    clean_data = {}
    if is_all:
        judge_list = [i for i in parameter_group_list if i[0] not in request_data and not i[2]]
        if judge_list:
            return False, parameter_error([i[0] for i in judge_list])
    for parameter, parameter_type, is_none, max_len in parameter_group_list:
        if parameter not in request_data and not is_all:
            continue
        value = request_data[parameter]
        if value == None:
            if not is_none:
                return False, parameter_error(f"{parameter}: None")
            else:
                continue
        if parameter_type == "file":
            suffix = max_len
            file_name = clean_file_name(value.filename.strip('"'))
            if not file_name.endswith(suffix):
                return False, parameter_error(f"{parameter}: {file_name}")
            clean_data[parameter] = (value, file_name)
            continue
        if not isinstance(value, parameter_type):
            if parameter_type == bool:
                try:
                    value = bool(int(value))
                except:
                    return False, parameter_error(f"{parameter}: {value}")
            try:
                value = parameter_type(value)
            except:
                return False, parameter_error(f"{parameter}: {value}")
        if len(str(value)) > max_len:
            return False, parameter_error(f"{parameter}: {value}")
        clean_data[parameter] = value
    return True, clean_data

def return_file(file_path, is_stream=False, delete_dir=False):
    base_dir, filename = os.path.split(file_path)
    if is_stream:
        def generate_stream():
            with open(file_path, "rb") as f:
                yield from f
            if delete_dir:
                file_dir = os.path.split(file_path)[0]
                shutil.rmtree(file_dir)
            else:
                os.remove(file_path)
        response = Response(generate_stream(), content_type='application/octet-stream')
    else:
        response = make_response(send_from_directory(base_dir, filename))
    response.headers["Content-Disposition"] = "attachment; filename={0}; filename*=utf-8''{0}".format(quote(filename))
    # print(type(response))
    return response

def generate_uuid_path(file_dir, file_name):
    file_suffix = os.path.splitext(file_name)[1]
    new_file_name = "".join((str(uuid.uuid1()), file_suffix))
    file_path = os.path.join(file_dir, new_file_name)
    return file_path

async def ocr_request_async(ocr_url, file_path):
    async with aiohttp.ClientSession() as session:
        file_name = os.path.split(file_path)[1]
        encode_data = encode_multipart_formdata(
            {"file": (file_name, open(file_path, 'rb').read()), "data": json.dumps({"caller_request_id": str(uuid.uuid1())})})
        data = encode_data[0]
        header = {'Content-Type': encode_data[1]}
        try:
            async with session.post(url=ocr_url, data=data, headers=header) as resp:
                # print(f"{file_path} start request")
                response_text = await resp.text()
                response_json = json.loads(response_text)
                img_data = [i["text_string"] for i in response_json["img_data_list"][0]["text_info"]]
        except Exception as e:
            print(f"{file_path} {e}")
            return False
        else:
            return img_data

def ocr_request(ocr_url, file_path):
    try:
        response = requests.post(ocr_url, files={"file": open(file_path, "rb")},
                                 data={"data": str(uuid.uuid1())})
        response_json = response.json()
        img_data = [i["text_string"] for i in response_json["img_data_list"][0]["text_info"]]
    except Exception as e:
        print(f"{file_path} {e}")
        return False
    else:
        return img_data

def extract(file_dir, file_name):
    file_path = os.path.join(file_dir, file_name)
    extract_dir = os.path.join(file_dir, "extract")
    if not os.path.exists(extract_dir):
        os.makedirs(extract_dir)
    if file_name.endswith((".zip", ".ZIP")):
        file = zipfile.ZipFile(file_path)
        for extract_file in file.namelist():
            file.extract(extract_file, extract_dir)
        file.close()
    if file_name.endswith((".rar", ".RAR")):
        file = rarfile.RarFile(file_path)
        file.extractall(extract_dir)
        file.close()
    os.remove(file_path)
    walk_list = os.walk(extract_dir)
    for pic_dir, _, pic_name_list in walk_list:
        for pic_name in pic_name_list:
            if pic_name.endswith(zip_suffix):
                extract(pic_dir, pic_name)

def extract_zip(base_dir, file_name, file):
    file_dir = os.path.join(base_dir, str(uuid.uuid1()))
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    file_path = os.path.join(file_dir, file_name)
    file.save(file_path)
    extract(file_dir, file_name)
    extract_dir = os.path.join(file_dir, "extract")
    walk_list = os.walk(extract_dir)
    return file_dir, extract_dir, walk_list

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell):
    kwargs = {
        "top": {"sz": 12, "val": "single"},
        "bottom": {"sz": 12, "val": "single"},
        "left": {"sz": 12, "val": "single"},
        "right": {"sz": 12, "val": "single"},
    }
    """
    Set cell`s border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        left={"sz": 24, "val": "dashed", "shadow": "true"},
        right={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('left', 'top', 'right', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))