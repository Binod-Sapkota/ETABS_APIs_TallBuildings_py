import sys
import os
import comtypes.client
import numpy as np

import pandas as pd
import json
from time import time
import NewModel_utils
import ObtainResults
import switcher as sw

import logging

logging.basicConfig(filename='test.log', level=logging.WARNING)

# =========== Program Settings =========================================================================
datafilename = "D:\\00_Thesis Project\ETABS_APIs_PYTHON\modelInputfiles\model.xlsx"
resultfilename = "D:\\00_Thesis Project\ETABS_APIs_PYTHON\excelfiles\\results.xlsx"
# AttachToInstance = False

runmode = "D"  # "I" - attach to instance, "D" - run from Directory, "N"- new model
startModel = 81  # start program from the startModel (index of list containing model directory)

nModels = 1

SpecifyPath = False
ProgramPath = "C:\Program Files\Computers and Structures\ETABS 18\ETABS.exe"
APIPath = 'E:\\00_Thesis Project\ETABS MODELS'

modelDirectory = "E:\\00_Thesis Project\\RUN_NOTMCE"
saveDirectory = "E:\\00_Thesis Project\\RUN_NOTMCE"


# ======================================================================================================

def runActiveModel():
    try:
        # get the active ETABS object
        myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)

    # create SapModel object
    SapModel = myETABSObject.SapModel
    # ret = SapModel.SetPresentUnits(10)

    resStoryfilename = "test2.xlsx"
    ObtainResults.obtainResults(SapModel, resStoryfilename)


def runFromDirectory():
    EtabsFile = NewModel_utils.ObjEtabsFile(modelDirectory)
    etabsfiles = EtabsFile.files
    fileroot = EtabsFile.roots
    nameBuilding = EtabsFile.names

    # -------------------- opening a ETABS Program --------------------------------------------------

    # create API helper object
    helper = comtypes.client.CreateObject('ETABSv1.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    if SpecifyPath:
        try:
            # 'create an instance of the ETABS object from the specified path
            myETABSObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program from " + ProgramPath)
            sys.exit(-1)
    else:

        try:
            # create an instance of the ETABS object from the latest installed ETABS
            myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program.")
            sys.exit(-1)

        # start ETABS application
    myETABSObject.ApplicationStart()

    # create SapModel object

    SapModel = myETABSObject.SapModel
    # ret = myETABSObject.Hide()

    nrow = startModel
    for f, rot, nam in zip(etabsfiles[(startModel - 1):len(etabsfiles)],
                           fileroot[(startModel - 1):len(etabsfiles)], nameBuilding[(startModel - 1):len(etabsfiles)]):
        nrow = nrow + 1

        dirs = os.path.basename(rot)
        savePath = os.path.join(saveDirectory, dirs)
        modelPath = savePath + os.sep + nam
        resPath = saveDirectory + os.sep + "storyResults"
        resStoryfilename = resPath + os.sep + nam.split(".")[0] + ".xlsx"

        ret = SapModel.File.OpenFile(f)



        # ========= if needed to save =====================================================
        if not os.path.exists(savePath):
            try:
                os.makedirs(savePath)
            except OSError:
                pass
        ret = SapModel.File.Save(modelPath)
        # ==================================================================================
        # input("Enter Any Key to Continue!!")  # stop program here (check ETABS file before processing)
        #
        #
        # if not os.path.exists(resPath):
        #     try:
        #         os.makedirs(resPath)
        #     except OSError:
        #         pass

        # Run for Results =========================================================================
        try:
            print("running >", nam)
            ObtainResults.obtainResults(SapModel, resStoryfilename, nrowGlobal=nrow, unitName=nam,resType="Global")
            print("completed >", nam)
        except:
            print("Not completed! skipped >", nam)

    myETABSObject.ApplicationExit(False)
    SapModel = None
    myETABSObject = None


# ======================================================================================================================

if runmode == "I":
    start_time = time()
    runActiveModel()
    end_time = time()
    print(f'Total time taken is {end_time - start_time} sec')
elif runmode == "D":
    runFromDirectory()
