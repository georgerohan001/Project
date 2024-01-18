import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import os
import shutil
import subprocess
import ast
import inspect
from importlib.machinery import SourceFileLoader


def extract_sheets(parent_folder):
    for folder_name in os.listdir(parent_folder):
        folder_path = os.path.join(parent_folder, folder_name)

        if os.path.isdir(folder_path):
            alr_ext = "Already Extracted Sheets"
            txt_files = "TXT Files"
            success_folder = "Successful Sheets"
            manual_folder = "Manual Correction Needed"
            unrecognized = "Unrecognized Sheets"

            already_extracted_folder = os.path.join(folder_path, alr_ext)
            os.makedirs(already_extracted_folder, exist_ok=True)

            txt_files_folder = os.path.join(folder_path, txt_files)
            os.makedirs(txt_files_folder, exist_ok=True)

            successful_sheets = os.path.join(folder_path, success_folder)
            os.makedirs(successful_sheets, exist_ok=True)

            manual_sheets = os.path.join(folder_path, manual_folder)
            os.makedirs(manual_sheets, exist_ok=True)

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


def sheet_mover(
    statement, points_log_path, script_path,
    folder_path, name, ex, task, points
):
    if statement:
        with open(points_log_path, 'r') as points_log_file:
            existing_content = points_log_file.read()
        previous_balance = int(
            existing_content.split("Point balance: ")[1].split("\n")[0])
        updated_point_balance = previous_balance + points
        updated_content = existing_content.replace(
            f"Point balance: {previous_balance}",
            f"Point balance: {updated_point_balance}")
        with open(points_log_path, 'w') as points_log:
            points_log.write(updated_content)
        with open(points_log_path, 'a') as points_log:
            points_log.write(
                f"Sheet {ex} Task {task} {name}: +{points} Points\n"
                )
        os.rename(script_path, os.path.join(
            folder_path, "Successful Sheets", name))
    elif os.path.exists(script_path):
        os.rename(script_path, os.path.join(
            folder_path, "Manual Correction Needed", name))


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
    statement = ((result.strip() == expected_output.strip()) and
                 (count_plus_equals <= 4)
                 )

    sheet_mover(statement, points_log_path, helloworld_script_path,
                folder_path, "helloworld.py", "01", "02", 6
                )


def username(folder_path):
    username_script_path = os.path.join(folder_path, 'username.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(
        ['python', username_script_path],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)

    output, _ = process.communicate(input='Barthelomew')

    statement = ('11' in output)

    sheet_mover(statement, points_log_path, username_script_path,
                folder_path, "username.py", "01", "03", 2
                )


def crosssum(folder_path):
    crosssum_script_path = os.path.join(folder_path, 'crosssum.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(['python', crosssum_script_path],
                               stdin=subprocess.PIPE,
                               stdout=subprocess.PIPE, text=True)
    output, _ = process.communicate(input='99')

    statement = ('18' in output)

    sheet_mover(statement, points_log_path, crosssum_script_path,
                folder_path, "crosssum.py", "01", "04", 2
                )


def lifeinweeks(folder_path):
    lifeinweeks_script_path = os.path.join(folder_path, 'lifeinweeks.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    process = subprocess.Popen(['python', lifeinweeks_script_path],
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                               text=True)

    output, _ = process.communicate(input='56')

    statement = ('12410 day' in output and
                 '1768 week' in output and
                 '408 month' in output)

    sheet_mover(statement, points_log_path, lifeinweeks_script_path,
                folder_path, "lifeinweeks.py", "01", "05", 5
                )


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

    statement = (successful_tests == 4)

    sheet_mover(statement, points_log_path, leapyear_script_path,
                folder_path, "leapyear.py", "02", "02", 4
                )


def million(folder_path):
    million_script_path = os.path.join(folder_path, 'million.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    result = subprocess.check_output(
        ['python', million_script_path], universal_newlines=True)

    statement = ("1" in result and
                 "1000000" in result and
                 "500000500000" in result)

    sheet_mover(statement, points_log_path, million_script_path,
                folder_path, "million.py", "02", "03", 2
                )


def caesar_cipher(folder_path):
    caesar_cipher_script_path = os.path.join(folder_path, 'caesar_cipher.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    result1 = subprocess.run(["python", caesar_cipher_script_path],
                             input="5\nZyzY", text=True, capture_output=True)
    result2 = subprocess.run(["python", caesar_cipher_script_path],
                             input="ZyzY\n5", text=True, capture_output=True)

    statement = (("EdeD" in result1.stdout) or ("EdeD" in result2.stdout))

    sheet_mover(statement, points_log_path, caesar_cipher_script_path,
                folder_path, "caesar_cipher.py", "02", "04", 9
                )


def books(folder_path):
    books_script_path = os.path.join(folder_path, 'books.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    books_module = SourceFileLoader('books', books_script_path).load_module()

    try:
        from books import Book, ChildrenBook, buy_books  # noqa: F401
    except ImportError:
        pass

    exercise_sheet_passed = True

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

    book_class = getattr(books_module, 'Book', None)

    book_class = getattr(books_module, 'Book', None)

    if book_class:
        signature = inspect.signature(book_class.__init__)

        required_attributes = ['title', 'author', 'price']
        present_attributes = all(
            param in signature.parameters for param in required_attributes)

        if present_attributes:
            pass
        else:
            exercise_sheet_passed = False

    if (
        hasattr(books_module, 'Book')
        and hasattr(books_module, 'ChildrenBook')
        and hasattr(Book, 'view') and callable(getattr(Book, 'view'))
        and (hasattr(ChildrenBook, 'view') and
             callable(getattr(ChildrenBook, 'view')))
        and hasattr(books_module, 'buy_books')
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
        var = locals()[f"{name}"]
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

    sheet_mover(exercise_sheet_passed, points_log_path, books_script_path,
                folder_path, "books.py", "03", "03", 4
                )


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

    sheet_mover(conditions_met, points_log_path, anagrams_script_path,
                folder_path, "anagrams.py", "03", "04", 3
                )


def data(folder_path):
    data_csv_path = os.path.join(folder_path, 'data.csv')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")

    statement = (os.path.exists(data_csv_path))

    sheet_mover(statement, points_log_path, data_csv_path,
                folder_path, "data.csv", "04", "01", 5
                )


def zen(folder_path):
    zen_script_path = os.path.join(folder_path, 'zen.py')
    points_log_path = os.path.join(folder_path, "Points_Log.txt")
    zen_imported = False

    with open(zen_script_path, 'r') as file:
        for line in file:
            if "import this" in line:
                zen_imported = True

    sheet_mover(zen_imported, points_log_path, zen_script_path,
                folder_path, "zen.py", "04", "03", 2
                )


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

    statement = (
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
    )

    sheet_mover(statement, points_log_path, shapes_script_path,
                folder_path, "shapes.py", "04", "04", 6
                )


def main():
    root = tk.Tk()
    root.withdraw()

    parent_folder = select_parent_folder()

    if parent_folder:
        extract_sheets(parent_folder)


if __name__ == "__main__":
    main()
