import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import os
import shutil
import subprocess
import unittest
import ast
from importlib.machinery import SourceFileLoader


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
            books(folder_path)
            anagrams(folder_path)
            data(folder_path)
            zen(folder_path)
            shapes(folder_path)

            folders_to_delete = ["__MACOSX", "__pycache__"]
            for folder_name in folders_to_delete:
                folder_to_delete_path = os.path.join(folder_path, folder_name)
                if (
                    os.path.exists(folder_to_delete_path)
                    and os.path.isdir(folder_to_delete_path)
                ):
                    shutil.rmtree(folder_to_delete_path)


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
    elif os.path.exists(helloworld_script_path):
        os.rename(helloworld_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "helloworld.py"))


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

    elif os.path.exists(username_script_path):
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

    elif os.path.exists(project_folder):
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
    elif os.path.exists(crosssum_script_path):
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
    elif os.path.exists(lifeinweeks_script_path):
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
    elif os.path.exists(leapyear_script_path):
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
    elif os.path.exists(million_script_path):
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
    elif os.path.exists(caesar_cipher_script_path):
        os.rename(caesar_cipher_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "caesar_cipher.py"))


def books(folder_path):
    books_script_path = os.path.join(folder_path, 'books.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    books_module = SourceFileLoader('books', books_script_path).load_module()

    try:
        from books import Book, ChildrenBook, buy_books  # noqa: F401
    except ImportError:
        pass

    classes = []
    class_methods = {}
    class_attributes = {}
    functions = []
    function_variables = {}
    exercise_sheet_passed = True

    def extract_info(node, current_class=None):
        if isinstance(node, ast.ClassDef):
            current_class = node.name
            classes.append(current_class)
            class_methods[current_class] = []
            class_attributes[current_class] = []
            for item in node.body:
                extract_info(item, current_class)
        elif isinstance(node, ast.FunctionDef):
            method_name = node.name
            if current_class:
                class_methods[current_class].append(method_name)
                for arg in node.args.args:
                    class_attributes[current_class].append(
                        f"{current_class}.{arg.arg}"
                    )
                for item in node.body:
                    extract_info(item, current_class)
            else:
                functions.append(method_name)
                function_variables[method_name] = []
                for item in node.body:
                    if isinstance(item, ast.Assign):
                        for target in item.targets:
                            if isinstance(target, ast.Name):
                                function_variables[method_name].append(
                                    target.id)
                    else:
                        extract_info(item)
        elif isinstance(node, ast.Assign):
            if current_class:
                for target in node.targets:
                    if isinstance(target, ast.Name):
                        variable_name = target.id
                        class_attributes[current_class].append(
                            f"{current_class}.{variable_name}"
                        )

        elif (
                isinstance(node, ast.Expr)
                and isinstance(node.value, ast.Call)
                and isinstance(node.value.func, ast.Name)
        ):
            function_name = node.value.func.id
            if current_class:
                class_methods[current_class].append(function_name)
        else:
            for item in ast.iter_child_nodes(node):
                extract_info(item, current_class)

    with open(books_script_path, 'r') as file:
        content = file.read()
        tree = ast.parse(content)

        for node in tree.body:
            extract_info(node)

    variable_classes = {}

    with open(books_script_path, 'r') as file:
        tree = ast.parse(file.read(), filename=books_script_path)

    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    variable_name = target.id
                    if isinstance(node.value, ast.Call):
                        value = ast.unparse(node.value)
                    else:
                        try:
                            value = ast.literal_eval(node.value)
                        except (ValueError, SyntaxError):
                            value = None
                    variable_classes[variable_name] = value

    variable_function_names = {}

    for variable_name, value in variable_classes.items():
        if isinstance(value, str) and value.startswith(
            ("Book(", "ChildrenBook(")
        ):
            class_name = value.split('(', 1)[0]
            variable_function_names[variable_name] = class_name
        else:
            variable_function_names[variable_name] = None

    book_count = 0
    childrenbook_count = 0
    for value in variable_function_names.values():
        if value == "Book":
            book_count += 1
        if value == "ChildrenBook":
            childrenbook_count += 1

    if (
        "Book" in classes
        and "ChildrenBook" in classes
        and "view" in class_methods["Book"]
        and "view" in class_methods["ChildrenBook"]
        and any(
            attr.lower() == "book.title" for attr in class_attributes["Book"])
        and any(
            attr.lower() == "book.author" for attr in class_attributes["Book"])
        and any(
            attr.lower() == "book.price" for attr in class_attributes["Book"])
        and "buy_books" in functions
        and book_count >= 3
        and childrenbook_count >= 2
    ):
        pass
    else:
        exercise_sheet_passed = False

    for variable in variable_classes:
        exec(f"{variable} = {variable_classes[variable]}")

    actual_sum = 0
    parameter_list = []
    for name in variable_function_names:
        var = locals()[name]
        if not (variable_function_names[name]):
            pass
        elif (
            (variable_function_names[name].lower() == "book")
            and (("Title" and "Author" and "Price") in var.view())
        ):
            pass
        elif (
            (variable_function_names[name].lower() == "childrenbook")
            and ("children's book" in var.view())
        ):
            pass
        else:
            exercise_sheet_passed = False

        if (
            variable_function_names[name]
            and (variable_function_names[name].lower() == "book")
        ):
            actual_sum += var.price
            parameter_list.append(name)

    if actual_sum == buy_books(*[locals()[param] for param in parameter_list]):
        pass
    else:
        exercise_sheet_passed = False

    if exercise_sheet_passed:
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
            points_log.write("Sheet 03 Task 03 books.py: +4 Points\n")

        os.rename(books_script_path, os.path.join(
            folder_path, "Successful Sheets", "books.py"))
    elif os.path.exists(books_script_path):
        os.rename(
            books_script_path,
            os.path.join(folder_path, "Unsuccessful Sheets", "books.py"))


def anagrams(folder_path):
    anagrams_script_path = os.path.join(folder_path, 'anagrams.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    result = subprocess.run(
        ["python", anagrams_script_path],
        capture_output=True, text=True,
        input="conversationalist\nconversationalist\n",
        check=True
        )

    conditions_met = True

    output = result.stdout.lower()

    if "anagram" in output and " not " not in output:
        pass
    else:
        conditions_met = False

    with open(anagrams_script_path, 'r') as file:
        code = file.read()

    try:
        parsed_code = ast.parse(code)
    except SyntaxError:
        pass

    dict_count = sum(
        isinstance(node, ast.Dict) for node in ast.walk(parsed_code)
        )

    if dict_count >= 1:
        pass
    else:
        conditions_met = False

    if conditions_met:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 3
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 03 Task 04 anagrams.py: +3 Points\n")

        os.rename(anagrams_script_path, os.path.join(
            folder_path, "Successful Sheets", "anagrams.py"))
    elif os.path.exists(anagrams_script_path):
        os.rename(anagrams_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "anagrams.py"))


def data(folder_path):
    data_csv_path = os.path.join(folder_path, 'data.csv')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    if os.path.exists(data_csv_path):
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
            points_log.write("Sheet 04 Task 01 data.csv: +5 Points\n")


def zen(folder_path):
    zen_script_path = os.path.join(folder_path, 'zen.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    zen_imported = False

    with open(zen_script_path, 'r') as file:
        for line in file:
            if "import this" in line:
                zen_imported = True

    if zen_imported:
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
            points_log.write("Sheet 04 Task 03 zen.py: +2 Points\n")

        os.rename(zen_script_path, os.path.join(
            folder_path, "Successful Sheets", "zen.py"))

    elif os.path.exists(zen_script_path):
        os.rename(zen_script_path, os.path.join(
            folder_path, "Unsuccessful Sheets", "zen.py"))


def shapes(folder_path):
    shapes_script_path = os.path.join(folder_path, 'shapes.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    modules_worked = True

    # Dynamically import the classes outside the loop
    shapes_module = (
        SourceFileLoader('shapes', shapes_script_path).load_module()
        )
    try:
        from shapes import Shape, Circle, Rectangle
    except ImportError:
        modules_worked = False
        print("one")

    # Instantiate the classes outside the loop
    try:
        shp = Shape(3)
        crcl = Circle(5)
        rct = Rectangle(3, 5)
    except TypeError:
        modules_worked = False

    try:
        bool(shp.area())
    except UnboundLocalError:
        modules_worked = False
    except TypeError:
        modules_worked = False

    if (
        modules_worked and
        hasattr(shapes_module, 'Shape') and
        hasattr(shapes_module, 'Circle') and
        hasattr(shapes_module, 'Rectangle') and
        issubclass(Circle, Shape) and
        issubclass(Rectangle, Shape) and
        (shp.area() and shp.perimeter()) and
        (78 <= float(crcl.area()) <= 79) and
        (31 <= float(crcl.perimeter()) <= 32) and
        (int(rct.area()) == 15) and
        (int(rct.perimeter()) == 16)
    ):
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()

        previous_balance = int(existing_content.split(
            "Point balance: ")[1].split("\n")[0])

        updated_point_balance = previous_balance + 6
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")

        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)

        with open(points_log_path, 'a') as points_log:
            points_log.write("Sheet 04 Task 04 shapes.py: +6 Points\n")

        os.rename(shapes_script_path, os.path.join(
            folder_path, "Successful Sheets", "shapes.py"))
    elif os.path.exists(shapes_script_path):
        os.rename(
            shapes_script_path,
            os.path.join(folder_path, "Unsuccessful Sheets", "shapes.py"))


def main():
    root = tk.Tk()
    root.withdraw()

    parent_folder = select_parent_folder()

    if parent_folder:
        extract_sheets(parent_folder)


if __name__ == "__main__":
    main()
