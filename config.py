import os

root_pass = os.path.dirname(os.path.abspath(__file__))


def get_path(*path):
    return os.path.join(root_pass, *path)