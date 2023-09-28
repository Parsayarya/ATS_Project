import os


def first_level_cleaning(dir_path: str, file_size_kb: int) -> None:
    """
    Delete files of a specific size in a given directory.

    :param dir_path: The path to the directory to search.
    :type dir_path: str
    :param file_size_kb: The file size in KB to search for.
    :type file_size_kb: int
    :raises FileNotFoundError: If the specified directory does not exist.
    :return: None
    """

    # Convert file size to bytes
    file_size_bytes = file_size_kb * 1000

    # Check if directory exists
    if not os.path.isdir(dir_path):
        raise FileNotFoundError(f"The specified directory does not exist: {dir_path}")

    # Loop through all files in the directory
    for filename in os.listdir(dir_path):
        # Get the full path to the file
        file_path = os.path.join(dir_path, filename)

        # Ensure that it's actually a file and not a directory
        if os.path.isfile(file_path):
            # Get the file size in bytes
            file_size_in_bytes = os.path.getsize(file_path)

            # Check if the file size matches the specified size
            if file_size_in_bytes == file_size_bytes:
                # Delete the file
                os.remove(file_path)
                print(f"Deleted: {file_path}")
            else:
                print(f"Skipped: {file_path} (Size: {file_size_in_bytes} bytes)")
        else:
            print(f"Skipped: {file_path} (Not a file)")


# first_level_cleaning(r'D:\python\ATS\DownloadCorpusBuilder\files\docxOutputs\IP',50.483)
# first_level_cleaning(r'D:\python\ATS\DownloadCorpusBuilder\files\docxOutputs\WP',50.483)
first_level_cleaning(r'D:\python\ATS\DownloadCorpusBuilder\files\docxOutputs\SIPN', 50.483)
first_level_cleaning(r'D:\python\ATS\DownloadCorpusBuilder\files\docxOutputs\SWPN', 50.483)
