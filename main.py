from exif import Image
from pathlib import Path
import shutil
import pendulum
import datetime
from win32com.propsys import propsys, pscon
import traceback


def uniquify(path: Path) -> Path:
    folder = path.parent
    stem = path.stem
    ext = path.suffix
    counter = 1

    while path.exists():
        new_filename = stem + " (" + str(counter) + ")" + ext
        path = folder.joinpath(new_filename)
        counter += 1

    return path


def ensure_directories(directory: Path) -> None:
    if not directory.exists():
        directory.mkdir(parents=True, exist_ok=True)


def handle_file(file: Path, target_directory: Path, should_move: bool) -> Path:
    if are_same_path(file.parent, target_directory):
        return file

    result = Path(shutil.copy2(str(file), str(target_directory)))

    if should_move:
        file.unlink()

    return result


def make_path(root: Path, date_time: datetime) -> Path:
    directory_schema = date_time.format("YYYY-MM")
    return root.joinpath(directory_schema)


def remove_prefix(text: str, prefix: str) -> str:
    if text.startswith(prefix):
        return text[len(prefix):]
    return text


def rename_file(file: Path, date_time: datetime, file_format_mode: int) -> Path:
    folder = file.parent
    stem = file.stem
    ext = file.suffix

    date_time = date_time.format("YYYY-MM-DD HH-mm-ss")

    # 1: date only
    # 2: preserve filenames
    # 3: date + old filename, do not duplicate date and remove excess whitespace
    # 4: date + old filename, do not alter old filename
    if file_format_mode is 1:
        new_filename = date_time + ext

    elif file_format_mode is 2:
        new_filename = stem + ext

    elif file_format_mode is 3:
        stem = remove_prefix(stem.strip(), date_time).strip()
        new_filename = date_time + " " + stem + ext

    elif file_format_mode is 4:
        new_filename = date_time + " " + stem + ext

    else:
        raise Exception("Filename format is invalid")

    new_file = folder.joinpath(new_filename)
    new_file = uniquify(new_file)
    file.rename(new_file)
    return file


def are_same_path(path1: Path, path2: Path) -> bool:
    if (not path1.exists()) | (not path2.exists()):
        return path1.absolute() is path2.absolute()
    else:
        return path1.samefile(path2)


def get_date_taken(file_path: Path) -> datetime:
    try:
        with open(str(file_path), 'rb') as file:
            my_image = Image(file)
            if my_image.has_exif:
                date_time = my_image.get('datetime_original')
                if date_time is not None:
                    return pendulum.from_format(date_time, "YYYY:MM:DD HH:mm:ss")
    except AssertionError:
        properties = propsys.SHGetPropertyStoreFromParsingName(str(file_path))
        dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()

        if dt is None:
            dt = properties.GetValue(pscon.PKEY_Media_DateReleased).GetValue()

        if dt is None:
            dt = properties.GetValue(pscon.PKEY_Photo_DateTaken).GetValue()

        if dt is None:
            dt = properties.GetValue(pscon.PKEY_RecordedTV_OriginalBroadcastDate).GetValue()

        if dt is not None:
            return pendulum.instance(dt)

    return None


def process_file(file_path: Path, target_root: Path, should_move: bool, file_format_mode: int) -> None:
    try:
        date_taken = get_date_taken(file_path)

        if date_taken is None:
            raise Exception("date_taken is None")

        print(date_taken)
        target_path = make_path(target_root, date_taken)
        ensure_directories(target_path)
        new_file = handle_file(file_path, target_path, should_move)
        new_file_renamed = rename_file(new_file, date_taken, file_format_mode)

        if are_same_path(file_path, new_file_renamed):
            print("Destination and source are the same file.")
        else:
            print(new_file_renamed)

    except Exception as ex:
        print("Error while parsing the date taken: " + str(ex))
        target_path = target_root.joinpath("Could not categorize")
        ensure_directories(target_path)
        new_file = handle_file(file_path, target_path, should_move)

        if are_same_path(file_path, new_file):
            print("Destination and source are the same file.")
        else:
            print(new_file)


def process_root(source_root: Path, target_root: Path, should_move: bool, file_format_mode: int) -> None:
    for filename in source_root.glob('**/*'):
        if filename.is_dir():
            print("directory")
            continue
        print(filename)
        process_file(filename, target_root, should_move, file_format_mode)
        print()


def main():
    source_root = Path(input("Source root: "))
    target_root = Path(input("Target root: "))
    move_or_copy = input("Move or copy? (m/c): ")
    file_format_mode = int(input(
        "Filename format of the new files? (1: date only / 2: preserve filenames / 3: date + old filename, "
        "do not duplicate date and remove excess whitespace / 4: date + old filename, do not alter old filename) ("
        "1/2/3/4): "))

    should_move = (move_or_copy is 'm') & (move_or_copy is not 'c')

    process_root(source_root, target_root, should_move, file_format_mode)


if __name__ == "__main__":
    main()
