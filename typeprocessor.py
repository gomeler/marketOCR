import yaml

import io
import pickle

class TypeProcessor(object):
    # TypeProcessor consumes the typeIDs.yaml from the static data export
    def __init__(self):
        # This should be a once-per-install process, so I'm not going to do too much work about making this very robust.
        self.yaml_data_file = "typeIDs.yaml"
        self.pickle_file = "typeIDs.pkl"
        self.typeIDs = None

    def load_data_from_yaml(self):
        # This is incredibly slow and will hang the process.
        with open(self.yaml_data_file, encoding="utf8", mode="r") as stream:
            data = yaml.load(stream)
        self.data = data

    def strip_data(self):
        # We really only care about the typeID and the name.
        stripped_data = {}
        for key, value in self.data.items():
            res = value.get("name").get("en")
            if res:
                stripped_data[key] = res
        self.stripped_data = stripped_data

    def save_ids_to_disk(self):
        # Dumping via pickle because it's the real dill.
        with open(self.pickle_file, mode="wb") as pickle_file:
            pickle.dump(self.stripped_data, pickle_file)

    def load_ids_from_disk(self):
        # Load the pickled id dict, and set it to self.typeIDs.
        try:
            with open(self.pickle_file, mode="rb") as pickle_file:
                self.typeIDs = pickle.load(pickle_file, encoding="utf8")
        except FileNotFoundError:
                raise FileNotFoundError(f"Pickle file {self.pickle_file} does not exist. "
                    "Perhaps you need to process typeIDs.yaml from the static data export?")

    # In the future this could hit up the API if a typeID isn't present in typeIDs.yaml
    def get(self, id):
        return self.typeIDs.get(int(id))