
import numpy as np
import os


class ObjStory:

    def __init__(self,name,areas):
        self.name = name
        self.areas = areas

    def get_area(self):
        return  np.sum(self.areas)



class ObjEtabsFile:
    def __init__(self, directoryName):
        self.files = []
        self.names = []
        self.dirs = []
        self.roots = []
        for root, dira, file in os.walk(directoryName):
            for filename in file:
                if filename.endswith(".EDB"):
                    self.names.append(filename)
                    self.roots.append(root)
                    self.files.append(os.path.join(root, filename))