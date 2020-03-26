import os
import yaml


def get_yaml(file_path):
    """
    Read Yaml file froma given file_path. If file does not exist, return None
    and don't try to open a file.
    """

    # If there is no such a file return from this function
    if not os.path.exists(file_path):
        print(f'Configuration file not found: {file_path}')
        return  # this ends a function execution immediately.

    # Open a file containing YAML format
    with open(file_path, 'r') as stream:
        try:
            # This will be parsed as a dictionary
            dat = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)
    return dat  # <- This can be used as dictionary
