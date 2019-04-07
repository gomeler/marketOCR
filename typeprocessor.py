import yaml

import io
import pickle

class TypeProcessor(object):
    # TypeProcessor consumes the typeIDs.yaml from the static data export
    def __init__(self):
        # This should be a once-per-install process, so I'm not going to do too much work about making this very robust.
        self.yaml_data_file = "typeIDs.yaml"

    def load_data(self):
        # This is incredibly slow and will hang the process.
        with open(self.yaml_data_file, encoding="utf8", mode="r") as stream:
            data = yaml.load(stream)
        self.data = data

    def strip_data(self):
        # We really only care about the typeID and the name.
        stripped_data = {}
        for key, value in self.data.items():
            stripped_data[key] = value.get('name').get('en')
        self.stripped_data = stripped_data

    def save_data_to_disk(self):
        # Dumping via pickle because it's the real dill.
        with open("typeIDs.pkl", "wb") as pickle_file:
            pickle.dump(self.stripped_data, pickle_file)