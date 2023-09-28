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
    if not os.path.exists(dir_path):
        raise FileNotFoundError(f"The specified directory does not exist: {dir_path}")

    # Iterate through the files in the specified directory
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            file_path = os.path.join(root, file)
            if os.path.getsize(file_path) == file_size_bytes:
                os.remove(file_path)

if __name__ == "__main__":
    # Example usage:
    # first_level_cleaning('directory_path/', 10)  # Delete all files of size 10 KB