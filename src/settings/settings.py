import pickle
from pathlib import Path


class Settings:
    NAMES1_FILE = Path("static/data/settings/names1.pkl")
    NAMES2_FILE = Path("static/data/settings/names2.pkl")
    REASON_FILE = Path("static/data/settings/reasons.pkl")

    # NAMES_FILE = Path("static/data/settings/names.pkl")
    def __init__(self):
        self._create_names1()
        self._create_names2()
        self._create_reason()

    @staticmethod
    def _create_data_files(path_file: str):
        with open(path_file, "wb") as f:
            new_list = []
            pickle.dump(new_list, f)

    @staticmethod
    def _check_file(path_file: Path) -> bool:
        return path_file.is_file()

    def _create_names1(self):
        if not self._check_file(self.NAMES1_FILE):
            self._create_data_files(str(self.NAMES1_FILE))

    def _create_names2(self):
        if not self._check_file(self.NAMES2_FILE):
            self._create_data_files(str(self.NAMES2_FILE))

    def _create_reason(self):
        if not self._check_file(self.REASON_FILE):
            self._create_data_files(str(self.REASON_FILE))


def save_data(data, filename):
    with open(filename, "wb") as f:
        pickle.dump(data, f)
    return list(data)


def load_data(filename):
    with open(filename, "rb") as f:
        return pickle.load(f)


def append_data(data, filename):
    old_data: list = load_data(filename)
    if data in old_data:
        old_data.remove(data)
    old_data.append(data)
    save_data(old_data, filename)
    return old_data


settings = Settings()
