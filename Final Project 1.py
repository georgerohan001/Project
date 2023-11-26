import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import os
import subprocess


def extract_sheets(parent_folder):
    for folder_name in os.listdir(parent_folder):
        folder_path = os.path.join(parent_folder, folder_name)

        if os.path.isdir(folder_path):
            already_extracted_folder = os.path.join(folder_path, "Already Extracted Sheets")
            os.makedirs(already_extracted_folder, exist_ok=True)

            txt_files_folder = os.path.join(folder_path, "TXT Files")
            os.makedirs(txt_files_folder, exist_ok=True)

            successful_sheets = os.path.join(folder_path, "Successful Sheets")
            os.makedirs(successful_sheets, exist_ok=True)

            unsuccessful_sheets = os.path.join(folder_path, "Unsuccessful Sheets")
            os.makedirs(unsuccessful_sheets, exist_ok=True)

            unrecognized_sheets = os.path.join(folder_path, "Unrecognized Sheets")
            os.makedirs(unrecognized_sheets, exist_ok=True)

            points_log_path = os.path.join(folder_path, "Points_Log.txt")
            if not os.path.exists(points_log_path):
                with open(points_log_path, 'w') as points_log:
                    point_balance = 5
                    points_log.write(f"File name: {folder_name}\n")
                    points_log.write(f"Point balance: {point_balance}\n")
                    points_log.write("\nLogs:\n")
                    points_log.write("Sheet 01 Task 01 IDE Installation: +5 Points\n")

            for sheet_name in ["sheet01.zip", "sheet02.zip", "sheet03.zip", "sheet04.zip", "sheet05.zip"]:
                sheet_zip_path = os.path.join(folder_path, sheet_name)

                if os.path.exists(sheet_zip_path):
                    if os.path.exists(os.path.join(already_extracted_folder, sheet_name)):
                        message = f"{sheet_name} has already been extracted for {folder_name}. Extraction for this file will be skipped."
                        messagebox.showinfo("Sheet Already Extracted", message)

                        not_extracted_path = os.path.join(folder_path, f"{sheet_name.replace('.zip', ' (Not extracted).zip')}")
                        os.rename(sheet_zip_path, not_extracted_path)

                        continue

                    with zipfile.ZipFile(sheet_zip_path, 'r') as zip_ref:
                        zip_ref.extractall(folder_path)

                    already_extracted_path = os.path.join(already_extracted_folder, sheet_name)
                    os.rename(sheet_zip_path, already_extracted_path)

                    for extracted_file in os.listdir(folder_path):
                        extracted_file_path = os.path.join(folder_path, extracted_file)

                        if extracted_file.lower().endswith('.txt') and extracted_file != "Points_Log.txt":
                            new_txt_path = os.path.join(txt_files_folder, f"{sheet_name.replace('.zip', '')}_{extracted_file}")
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
        result = subprocess.check_output(['python', helloworld_script_path], universal_newlines=True)

        expected_output = "=========================\n\nH+e+l+l+o W+o+r+l+d !!!\n\n========================="

        if result.strip() == expected_output.strip() and count_plus_equals <= 4:
            with open(points_log_path, 'r') as points_log_file:
                existing_content = points_log_file.read()

            previous_balance = int(existing_content.split("Point balance: ")[1].split("\n")[0])

            updated_point_balance = previous_balance + 6
            updated_content = existing_content.replace(f"Point balance: {previous_balance}", f"Point balance: {updated_point_balance}")

            with open(points_log_path, 'w') as points_log:
                points_log.write(updated_content)

            with open(points_log_path, 'a') as points_log:
                points_log.write("Sheet 01 Task 02 helloworld.py: +6 Points\n")

            os.rename(helloworld_script_path, os.path.join(folder_path, "Successful Sheets", "helloworld.py"))
        else:
            os.rename(helloworld_script_path, os.path.join(folder_path, "Unsuccessful Sheets", "helloworld.py"))

    except subprocess.CalledProcessError as e:
        print("Error: ", e.output)


def username(folder_path):
    username_script_path = os.path.join(folder_path, 'username.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(['python', username_script_path], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
    output, _ = process.communicate(input='Barthelomew')

    if '11' in output:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split("Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 2
        updated_content = existing_content.replace(f"Point balance: {previous_balance}", f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 01 Task 03 username.py: +2 Points\n")

        os.rename(username_script_path, os.path.join(folder_path, "Successful Sheets", "username.py"))
    else:
        os.rename(username_script_path, os.path.join(folder_path, "Unsuccessful Sheets", "username.py"))


def crosssum(folder_path):
    crosssum_script_path = os.path.join(folder_path, 'crosssum.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(['python', crosssum_script_path], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
    output, _ = process.communicate(input='99')

    if '18' in output:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split("Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 2
        updated_content = existing_content.replace(f"Point balance: {previous_balance}", f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 01 Task 04 crosssum.py: +2 Points\n")

        os.rename(crosssum_script_path, os.path.join(folder_path, "Successful Sheets", "crosssum.py"))
    else:
        os.rename(crosssum_script_path, os.path.join(folder_path, "Unsuccessful Sheets", "crosssum.py"))


def lifeinweeks(folder_path):
    lifeinweeks_script_path = os.path.join(folder_path, 'lifeinweeks.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(['python', lifeinweeks_script_path], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
    output, _ = process.communicate(input='56')

    if ('12410 day' in output) and ('1768 week' in output) and ('408 month' in output):
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split("Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 5
        updated_content = existing_content.replace(f"Point balance: {previous_balance}", f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 01 Task 05 lifeinweeks.py: +5 Points\n")

        os.rename(lifeinweeks_script_path, os.path.join(folder_path, "Successful Sheets", "lifeinweeks.py"))
    else:
        os.rename(lifeinweeks_script_path, os.path.join(folder_path, "Unsuccessful Sheets", "lifeinweeks.py"))


def leapyear(folder_path):
    test_cases = [2000, 2004, 2100, 2097]
    successful_tests = 0

    for test_case in test_cases:
        leapyear_script_path = os.path.join(folder_path, 'leapyear.py')
        points_log_path = os.path.join(folder_path, "Points_Log.txt")

        process = subprocess.Popen(['python', leapyear_script_path], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
        output, _ = process.communicate(input=str(test_case))

        if (test_case == 2000 and 'Leap year' in output and 'Not' not in output) or \
           (test_case == 2004 and 'Leap year' in output and 'Not' not in output) or \
           (test_case == 2100 and 'Not' in output) or \
           (test_case == 2097 and 'Not' in output):
            successful_tests += 1

    if successful_tests == 4:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split("Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 4
        updated_content = existing_content.replace(f"Point balance: {previous_balance}", f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 02 Task 02 leapyear.py: +4 Points\n")

        os.rename(leapyear_script_path, os.path.join(folder_path, "Successful Sheets", "leapyear.py"))
    else:
        os.rename(leapyear_script_path, os.path.join(folder_path, "Unsuccessful Sheets", "leapyear.py"))


def million(folder_path):
    million_script_path = os.path.join(folder_path, 'million.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    result = subprocess.check_output(['python', million_script_path], universal_newlines=True)

    if ("1" in result) and ("1000000" in result) and "500000500000" in result:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split("Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 2
        updated_content = existing_content.replace(f"Point balance: {previous_balance}", f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 02 Task 03 million.py: +2 Points\n")

        os.rename(million_script_path, os.path.join(folder_path, "Successful Sheets", "million.py"))
    else:
        os.rename(million_script_path, os.path.join(folder_path, "Unsuccessful Sheets", "million.py"))


def caesar_cipher(folder_path):
    caesar_cipher_script_path = os.path.join(folder_path, 'caesar_cipher.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    result1 = subprocess.run(["python", caesar_cipher_script_path], input="5\nZyzY", text=True, capture_output=True)
    result2 = subprocess.run(["python", caesar_cipher_script_path], input="ZyzY\n5", text=True, capture_output=True)

    if ("EdeD" in result1.stdout) or ("EdeD" in result2.stdout):
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split("Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 9
        updated_content = existing_content.replace(f"Point balance: {previous_balance}", f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 02 Task 04 caesar_cipher.py: +9 Points\n")

        os.rename(caesar_cipher_script_path, os.path.join(folder_path, "Successful Sheets", "caesar_cipher.py"))
    else:
        os.rename(caesar_cipher_script_path, os.path.join(folder_path, "Unsuccessful Sheets", "caesar_cipher.py"))


def main():
    root = tk.Tk()
    root.withdraw()

    parent_folder = select_parent_folder()

    if parent_folder:
        extract_sheets(parent_folder)


if __name__ == "__main__":
    main()
