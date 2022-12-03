import os
import re
import xlsxwriter

file_path = "./log/"
excel_file_name = "HF.xlsx"


def get_hf_and_dipole_value_from_line(line):
    pattern = r'HF=(-?\d+\.?\d*).*Dipole=(-?\d+\.?\d*),(-?\d+\.?\d*),(-?\d+\.?\d*)'
    searchObj = re.search(pattern, line)
    if searchObj:
        return searchObj.group(1), (searchObj.group(2), searchObj.group(3), searchObj.group(4))


def get_hf_and_dipole_text_line(f):
    filename = os.path.join(file_path, f)
    blocks = []
    with open(filename, 'r') as f:
        in_block_flag = False
        line = f.readline()
        while line:
            if "GINC" in line:
                in_block_flag = True
            if "The archive entry for this job was punched" in line:
                in_block_flag = False
            if in_block_flag:
                blocks.append(line.strip())
            line = f.readline()
    return ''.join(blocks).replace('\r', '').replace('\n', '')


def get_hf_and_dipole_from_file(f):
    block = get_hf_and_dipole_text_line(f)
    return get_hf_and_dipole_value_from_line(block)


def get_all_hf_and_dipole_value_from_folder(path):
    files = os.listdir(path)
    values = []
    for f in files:
        if os.path.isfile(os.path.join(path, f)):
            hf, dipole = get_hf_and_dipole_from_file(f)
            fn = os.path.splitext(f)[0]
            values.append([fn, hf, dipole])
    return values


def dump_hf_and_dipole_values_to_excel(file_name, values):
    workbook = xlsxwriter.Workbook(file_name, {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet(file_name)

    row = 0
    header_format = workbook.add_format({'bold': True})
    header_format.set_align('center')
    worksheet.write_string(row, 0, "File Name", header_format)
    worksheet.write_string(row, 1, "HF", header_format)
    worksheet.merge_range(row, 2, row, 4, "Dipole", header_format)

    row += 1
    for filename, hf_value, (dipole1_value, dipole2_value, dipole3_value) in values:
        worksheet.write_string(row, 0, filename)
        worksheet.write(row, 1, hf_value)
        worksheet.write(row, 2, dipole1_value)
        worksheet.write(row, 3, dipole2_value)
        worksheet.write(row, 4, dipole3_value)
        row += 1
    workbook.close()


if __name__ == '__main__':
    hf_dipole_values = get_all_hf_and_dipole_value_from_folder(file_path)
    dump_hf_and_dipole_values_to_excel(excel_file_name, hf_dipole_values)
