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


def copy_file(file, target_directory):
    return shutil.copy(file, target_directory)


def make_path(root, datetime):
    directory_schema = datetime.format("YYYY-MM")
    result = os.path.join(root, directory_schema)
    return Path(result)


def rename_file(file, datetime):
    folder = Path(file).parent
    filename, extension = os.path.splitext(file)
    new_filename = datetime.format("YYYY-MM-DD HH-mm-ss") + extension
    new_file = os.path.join(folder, new_filename)
    new_file = uniquify(new_file)
    os.rename(file, new_file)
    return new_file


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


def process_file(file_path, target_root):
    try:
        date_taken = get_date_taken(file_path)
        print(date_taken)
        target_path = make_path(target_root, date_taken)
        ensure_directories(target_path)
        new_file = copy_file(file_path, target_path)
        new_file_renamed = rename_file(new_file, date_taken)
        print(new_file_renamed)
    except Exception as ex:
        print("Error while parsing the date taken: " + str(ex))
        target_path = os.path.join(target_root, "Could not categorize")
        ensure_directories(target_path)
        new_file = copy_file(file_path, target_path)
        print(new_file)


def process_root(source_root, target_root):
    for filename in Path(source_root).glob('**/*'):
        if filename.is_dir():
            print("directory")
            continue
        print(filename)
        process_file(filename, target_root)
        print()


def main():
    source_root = input("Source root: ")
    target_root = input("Target root: ")
    process_root(source_root, target_root)


if __name__ == "__main__":
    main()
