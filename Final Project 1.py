import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import os
import subprocess
import unittest


def extract_sheets(parent_folder):
    for folder_name in os.listdir(parent_folder):
        folder_path = os.path.join(parent_folder, folder_name)

        if os.path.isdir(folder_path):
            alr_ext = "Already Extracted Sheets"
            txt_files = "TXT Files"
            success_folder = "Successful Sheets"
            unsuccess_folder = "Unsuccessful Sheets"
            unrecognized = "Unrecognized Sheets"

            already_extracted_folder = os.path.join(folder_path, alr_ext)
            os.makedirs(already_extracted_folder, exist_ok=True)

            txt_files_folder = os.path.join(folder_path, txt_files)
            os.makedirs(txt_files_folder, exist_ok=True)

            successful_sheets = os.path.join(folder_path, success_folder)
            os.makedirs(successful_sheets, exist_ok=True)

            unsuccessful_sheets = os.path.join(folder_path, unsuccess_folder)
            os.makedirs(unsuccessful_sheets, exist_ok=True)

            unrecognized_sheets = os.path.join(folder_path, unrecognized)
            os.makedirs(unrecognized_sheets, exist_ok=True)

            points_log_path = os.path.join(folder_path, "Points_Log.txt")
            if not os.path.exists(points_log_path):
                with open(points_log_path, 'w') as points_log:
                    point_balance = 5
                    ex01_log = "Sheet 01 Task 01 IDE Installation: +5 Points\n"
                    points_log.write(f"File name: {folder_name}\n")
                    points_log.write(f"Point balance: {point_balance}\n")
                    points_log.write("\nLogs:\n")
                    points_log.write(ex01_log)

            sheet_names = ["sheet01.zip",
                           "sheet02.zip",
                           "sheet03.zip",
                           "sheet04.zip",
                           "sheet05.zip"]

            for sheet_name in sheet_names:
                sheet_zip_path = os.path.join(folder_path, sheet_name)

                if os.path.exists(sheet_zip_path):
                    if os.path.exists(
                        os.path.join(already_extracted_folder, sheet_name)
                    ):
                        message = (
                            f"{sheet_name} has already been extracted for"
                            f"{folder_name}.Extraction for this file "
                            "will be skipped."
                        )
                        messagebox.showinfo("Sheet Already Extracted", message)

                        zip_files = sheet_name.replace(
                            '.zip', ' (Not extracted).zip')

                        not_extracted_path = os.path.join(
                            folder_path, f"{zip_files}")

                        os.rename(sheet_zip_path, not_extracted_path)

                        continue

                    with zipfile.ZipFile(sheet_zip_path, 'r') as zip_ref:
                        zip_ref.extractall(folder_path)

                    # Move elements out of additional folder, if present
                    additional_folder = sheet_name.replace('.zip', '')
                    additional_folder_path = os.path.join(folder_path,
                                                          additional_folder)

                    if os.path.exists(additional_folder_path) and (
                            os.path.isdir(additional_folder_path)):
                        for item in os.listdir(additional_folder_path):
                            item_path = os.path.join(additional_folder_path,
                                                     item)
                            new_item_path = os.path.join(folder_path, item)
                            os.rename(item_path, new_item_path)

                        # Remove the now empty additional folder
                        os.rmdir(additional_folder_path)

                    already_extracted_path = os.path.join(
                        already_extracted_folder, sheet_name)
                    os.rename(sheet_zip_path, already_extracted_path)

                    for extracted_file in os.listdir(folder_path):
                        extracted_file_path = os.path.join(
                            folder_path, extracted_file)

                        if extracted_file.lower().endswith(
                                '.txt') and extracted_file != "Points_Log.txt":
                            new_txt_path = os.path.join(
                                txt_files_folder,
                                f"{sheet_name.replace('.zip', '')}"
                                f"_{extracted_file}")

                            os.rename(extracted_file_path, new_txt_path)

            helloworld(folder_path)
            username(folder_path)
            crosssum(folder_path)
            lifeinweeks(folder_path)
            leapyear(folder_path)
            million(folder_path)
            caesar_cipher(folder_path)


def select_parent_folder():
    parent_folder = filedialog.askdirectory(title="Select Parent Folder")
    return parent_folder


def helloworld(folder_path):
    helloworld_script_path = os.path.join(folder_path, 'helloworld.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    count_plus_equals = 0

    with open(helloworld_script_path, 'r') as file:
        for line in file:
            if line.count('print') > 0:
                count_plus_equals += line.count("+") + line.count("=")

    try:
        result = subprocess.check_output(
            ['python', helloworld_script_path], universal_newlines=True
        )

        expected_output = f"{25*'='}\n\nH+e+l+l+o W+o+r+l+d !!!\n\n{25*'='}"

        if (
            result.strip() == expected_output.strip()) and (
                count_plus_equals <= 4):

            with open(points_log_path, 'r') as points_log_file:
                existing_content = points_log_file.read()

            previous_balance = int(
                existing_content.split("Point balance: ")[1].split("\n")[0])

            updated_point_balance = previous_balance + 6
            updated_content = existing_content.replace(
                f"Point balance: {previous_balance}",
                f"Point balance: {updated_point_balance}")

            with open(points_log_path, 'w') as points_log:
                points_log.write(updated_content)

            with open(points_log_path, 'a') as points_log:
                points_log.write("Sheet 01 Task 02 helloworld.py: +6 Points\n")

            os.rename(helloworld_script_path, os.path.join(
                folder_path, "Successful Sheets", "helloworld.py"))
        else:
            os.rename(helloworld_script_path, os.path.join(
                folder_path, "Unsuccessful Sheets", "helloworld.py"))

    except subprocess.CalledProcessError as e:
        print("Error: ", e.output)


def username(folder_path):
    username_script_path = os.path.join(folder_path, 'username.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(
        ['python', username_script_path],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)

    output, _ = process.communicate(input='Barthelomew')

    if '11' in output:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 2
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 01 Task 03 username.py: +2 Points\n")

        os.rename(username_script_path, os.path.join(
            folder_path, "Successful Sheets", "username.py"))
    else:
        os.rename(
            username_script_path,
            os.path.join(folder_path, "Unsuccessful Sheets", "username.py"))


def project(folder_path):
    project_zip_path = os.path.join(folder_path, 'project.zip')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    code_folder = os.path.join(folder_path, 'code')
    test_file = os.path.join(folder_path, 'test.py')
    project_folder = os.path.join(folder_path, 'project')

    with zipfile.ZipFile(project_zip_path, 'r') as zip_ref:
        zip_ref.extractall(folder_path)

    if os.path.exists(code_folder) and os.path.exists(test_file):
        os.makedirs(project_folder, exist_ok=True)

        os.rename(code_folder, os.path.join(project_folder, 'code'))
        os.rename(test_file, os.path.join(project_folder, 'test.py'))

        code_folder = os.path.join(project_folder, 'code')
        test_file = os.path.join(project_folder, 'test.py')

    else:
        code_folder = os.path.join(project_folder, 'code')
        test_file = os.path.join(project_folder, 'test.py')

    loader = unittest.TestLoader()
    suite = loader.loadTestsFromName(test_file)
    num_tests = suite.countTestCases()

    try:
        with open(test_file, 'r') as file:
            has_required_imports = any(line.strip().startswith(
                'import code.caesercipher') or line.strip().startswith(
                    'import code.caesarcipher') or line.strip().startswith(
                        'from code.caesercipher import') or
                    line.strip().startswith(
                        'from code.caesarcipher import') for line in file)

    except FileNotFoundError:
        has_required_imports = False

    with open(os.devnull, 'w') as null_stream:
        result = unittest.TextTestRunner(stream=null_stream).run(suite)
    all_tests_passed = result.wasSuccessful()

    if all_tests_passed and has_required_imports and (num_tests > 2):
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 10
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 03 Task 01 project.zip: +10 Points\n")

        os.rename(project_folder, os.path.join(
            folder_path, "Successful Sheets", "project"))

    else:
        os.rename(project_folder, os.path.join(
            folder_path, "Unsuccessful Sheets", "project"))


def crosssum(folder_path):
    crosssum_script_path = os.path.join(folder_path, 'crosssum.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(['python', crosssum_script_path],
                               stdin=subprocess.PIPE,
                               stdout=subprocess.PIPE, text=True)
    output, _ = process.communicate(input='99')

    if '18' in output:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 2
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 01 Task 04 crosssum.py: +2 Points\n")

        os.rename(crosssum_script_path, os.path.join(
            folder_path, "Successful Sheets", "crosssum.py"))
    else:
        os.rename(crosssum_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "crosssum.py"))


def lifeinweeks(folder_path):
    lifeinweeks_script_path = os.path.join(folder_path, 'lifeinweeks.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(['python', lifeinweeks_script_path],
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                               text=True)

    output, _ = process.communicate(input='56')

    if (
        '12410 day' in output) and (
            '1768 week' in output) and (
                '408 month' in output):

        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 5
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 01 Task 05 lifeinweeks.py: +5 Points\n")

        os.rename(lifeinweeks_script_path, os.path.join(
            folder_path, "Successful Sheets", "lifeinweeks.py"))
    else:
        os.rename(lifeinweeks_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "lifeinweeks.py"))


def leapyear(folder_path):
    test_cases = [2000, 2004, 2100, 2097]
    successful_tests = 0

    for test_case in test_cases:
        leapyear_script_path = os.path.join(folder_path, 'leapyear.py')
        points_log_path = os.path.join(folder_path, "Points_Log.txt")

        process = subprocess.Popen(['python', leapyear_script_path],
                                   stdin=subprocess.PIPE,
                                   stdout=subprocess.PIPE,
                                   text=True)

        output, _ = process.communicate(input=str(test_case))

        if ((
            test_case == 2000) and (
                'Leap year' in output) and (
                    'Not' not in output)) or \
           ((
               test_case == 2004) and (
                   'Leap year' in output) and (
                       'Not' not in output)) or \
           ((
               test_case == 2100) and (
                   'Not' in output)) or \
           ((
               test_case == 2097) and (
                   'Not' in output)):

            successful_tests += 1

    if successful_tests == 4:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 4
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 02 Task 02 leapyear.py: +4 Points\n")

        os.rename(leapyear_script_path, os.path.join(
            folder_path, "Successful Sheets", "leapyear.py"))
    else:
        os.rename(leapyear_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "leapyear.py"))


def million(folder_path):
    million_script_path = os.path.join(folder_path, 'million.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    result = subprocess.check_output(
        ['python', million_script_path], universal_newlines=True)

    if ("1" in result) and ("1000000" in result) and "500000500000" in result:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 2
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 02 Task 03 million.py: +2 Points\n")

        os.rename(million_script_path, os.path.join(
            folder_path, "Successful Sheets", "million.py"))
    else:
        os.rename(million_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "million.py"))


def caesar_cipher(folder_path):
    caesar_cipher_script_path = os.path.join(folder_path, 'caesar_cipher.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    result1 = subprocess.run(["python", caesar_cipher_script_path],
                             input="5\nZyzY", text=True, capture_output=True)
    result2 = subprocess.run(["python", caesar_cipher_script_path],
                             input="ZyzY\n5", text=True, capture_output=True)

    if ("EdeD" in result1.stdout) or ("EdeD" in result2.stdout):
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 9
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 02 Task 04 caesar_cipher.py: +9 Points\n")

        os.rename(caesar_cipher_script_path, os.path.join(
            folder_path, "Successful Sheets", "caesar_cipher.py"))
    else:
        os.rename(caesar_cipher_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "caesar_cipher.py"))


def main():
    root = tk.Tk()
    root.withdraw()

    parent_folder = select_parent_folder()

    if parent_folder:
        extract_sheets(parent_folder)


if __name__ == "__main__":
    main()
