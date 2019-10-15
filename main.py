from exif import Image
from pathlib import Path
import shutil
import pendulum
import os
import pytz
import datetime
from win32com.propsys import propsys, pscon


def uniquify(path):
    filename, extension = os.path.splitext(path)
    counter = 1

    while os.path.exists(path):
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1

    return path


def ensure_directories(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


def handle_file(file, target_directory, shouldMove):
    str_file = str(Path(file).resolve())
    str_target_directory = str(Path(target_directory).resolve())

    try:
        str_result = shutil.copy(str_file, str_target_directory)
    except shutil.SameFileError:
        return str_file
    else:
        if shouldMove:
            os.remove(file)
        return str_result
        

def make_path(root, datetime):
    directory_schema = datetime.format("YYYY-MM")
    result = os.path.join(root, directory_schema)
    return Path(result)


def remove_prefix(text, prefix):
    if text.startswith(prefix):
        return text[len(prefix):]
    return text  # or whatever


def rename_file(file, datetime, file_format_mode):
    folder = Path(file).parent
    filename, extension = os.path.splitext(file)

    datetime = datetime.format("YYYY-MM-DD HH-mm-ss")

    # 1: date only
    # 2: preserve filenames
    # 3: date + old filename, do not duplicate date and remove excess whitespace
    # 4: date + old filename, do not alter old filename
    if file_format_mode is 1:
        new_filename = datetime + extension
    
    elif file_format_mode is 2:
        new_filename = filename + extension
    
    elif file_format_mode is 3:
        filename = remove_prefix(filename.strip(), datetime).strip()
        new_filename = datetime + " " + filename + extension
    
    elif file_format_mode is 4:
        new_filename = datetime + " " + filename + extension
    
    else:
        raise Exception("Filename format is invalid")

    new_file = os.path.join(folder, new_filename)
    new_file = uniquify(new_file)
    os.rename(file, new_file)
    return new_file


def are_same_file(path1, path2):
    return Path(path1).resolve().samefile(Path(path2).resolve())


def get_date_taken(file_path):
    try:
        with open(file_path, 'rb') as file:
            my_image = Image(file)
            if my_image.has_exif:
                date_time = my_image.get('datetime_original')
                if date_time is not None:
                    return pendulum.from_format(date_time, "YYYY:MM:DD HH:mm:ss")
    except:
        properties = propsys.SHGetPropertyStoreFromParsingName(str(Path(file_path).absolute()))
        dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()

        if dt is None:
            dt = properties.GetValue(pscon.PKEY_Media_DateReleased).GetValue()

        if dt is None:
            dt = properties.GetValue(pscon.PKEY_Photo_DateTaken).GetValue()

        if dt is None:
            dt = properties.GetValue(pscon.PKEY_RecordedTV_OriginalBroadcastDate).GetValue()

        if dt is not None:
            if not isinstance(dt, datetime.datetime):
                dt = datetime.datetime.fromtimestamp(int(dt))

            return pendulum.from_timestamp(dt.timestamp())


def process_file(file_path, target_root, shouldMove, file_format_mode):
    try:
        date_taken = get_date_taken(file_path)
        
        if date_taken is None:
            raise Exception("date_taken is None")

        print(date_taken)
        target_path = make_path(target_root, date_taken)
        ensure_directories(target_path)
        new_file = handle_file(file_path, target_path, shouldMove)
        new_file_renamed = rename_file(new_file, date_taken, file_format_mode)
        
        if are_same_file(file_path, new_file_renamed):
            print("Destination and source are the same file.")
        else:
            print(new_file_renamed)

    except Exception as ex:
        print("Error while parsing the date taken: " + str(ex))
        target_path = os.path.join(target_root, "Could not categorize")
        ensure_directories(target_path)
        new_file = handle_file(file_path, target_path, shouldMove)

        if are_same_file(file_path, new_file):
            print("Destination and source are the same file.")
        else:
            print(new_file)


def process_root(source_root, target_root, shouldMove, file_format_mode):
    for filename in Path(source_root).glob('**/*'):
        if filename.is_dir():
            print("directory")
            continue
        print(filename)
        process_file(filename, target_root, shouldMove, file_format_mode)
        print()


def main():
    source_root = input("Source root: ")
    target_root = input("Target root: ")
    move_or_copy = input("Move or copy? (m/c): ")
    file_format_mode = int(input("Filename format of the new files? (1: date only / 2: preserve filenames / 3: date + old filename, do not duplicate date and remove excess whitespace / 4: date + old filename, do not alter old filename) (1/2/3/4): "))

    shouldMove = (move_or_copy is 'm') & (move_or_copy is not 'c')

    process_root(source_root, target_root, shouldMove, file_format_mode)


if __name__ == "__main__":
    main()
