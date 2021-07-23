import sys

import comtypes.client
# import comtypes.gen
import numpy as np
import os
import pandas as pd
import json

import Obj_utils
import ObtainResults
import OAPIs

# =========== Program Settings =========================================================================
datafilename = "model.xlsx"
AttachToInstance = True
SpecifyPath = False
ProgramPath = "C:\Program Files\Computers and Structures\ETABS 18\ETABS.exe"
APIPath = 'E:\\20-10-14_ETABS_OAPIs_PYTHON\ETABS MODELS'


# ======================================================================================================


class StructureModel:

    def __init__(self):
        pass

    def set_modelInfo(self, modelInfo):
        pass

    def createModel(self, modelInfo, sectionInfo, modelName="MODEL_001.edb", **kwargs):

        if not os.path.exists(APIPath):
            try:
                os.makedirs(APIPath)
            except OSError:
                pass
        ModelPath = APIPath + os.sep + modelName

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

        # ================ Draw Model ===============================
        topRightCornerJoint = Obj_utils.drawNewModel(SapModel, ModelPath, datafilename,
                                                     modelInfo, sectionInfo)

        # ********************** Test Code ***********************************************
        # from test_Obj_utils import testdatabaseTable
        # NumberTables, TableKey, TableName, ImportType = testdatabaseTable(SapModel)

        LTYPE_QUAKE = 5
        ret = SapModel.LoadPatterns.Add('EQ', LTYPE_QUAKE, 0, True)

        tTableKey = 'Story Definitions'
        GroupName = ""
        [fieldKeyList, tableversion, fieldkeyincluded, numberrecords, tabledata,
         ret] = SapModel.DatabaseTables.GetTableForDisplayArray(TableKey=tTableKey, FieldKeyList="", GroupName="")

        data = np.asarray(tabledata).reshape((numberrecords, len(fieldkeyincluded)))
        # *********************************************************************************************************

        # ============== Analysis Setup ===============================
        Obj_utils.analysisSetup(SapModel)

        # ============= Obtain Joint Results =========================
        [U1, roofDrift] = ObtainResults.getJointResults(SapModel, jointName=topRightCornerJoint)

        # ============= Obtain story column Info =========================
        columns = ObtainResults.getStoryColumnInfo(SapModel, story="Story1")

        # ============= Obtain Bending Rigidity Index (BRI) =========================
        BRI = ObtainResults.getBRI(SapModel, n_corners=2)

        myETABSObject.ApplicationExit(False)
        SapModel = None
        myETABSObject = None

        return np.array([U1, roofDrift, BRI], ndmin=2)

    def runActiveModel(self):
        try:
            # get the active ETABS object
            myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        except (OSError, comtypes.COMError):
            print("No running instance of the program found or failed to attach.")
            sys.exit(-1)

        # create SapModel object
        SapModel = myETABSObject.SapModel

        # Obj_utils.defineAutoSeismic_UserCoefficient(SapModel)

        # dta = ObtainResults.get_ConcreteProperty(SapModel)
        #
        # ans = ObtainResults.get_dataFromTable(SapModel, tableKey = "Story Drifts",
        #                                       fieldKeyList =['OutputCase','Story','Drift'],
        #                                       fieldKeyTypeList=[np.str,np.str,np.float],
        #                                       required_column=['Story','Drift'], filterKeyList=['OutputCase'],
        #                                       required_row=['EQ'])

        # ans2 = ObtainResults.get_dataFromTable(SapModel, tableKey="Story Max Over Avg Displacements",
        #                                        fieldKeyList=['Story', 'OutputCase','Direction', 'Maximum'],
        #                                        fieldKeyTypeList=[np.str, np.str, np.float],
        #                                        required_column=['Story', 'Maximum'],
        #                                        filterKeyList=['OutputCase','Direction'],
        #                                        required_row=[['EQ', 'LATERAL'],['X']])

        # ans2 = ObtainResults.get_dataFromTable(SapModel, tableKey="Story Max Over Avg Displacements",
        #                                        fieldKeyList=['Story', 'OutputCase','Direction', 'Maximum'],
        #                                        required_column=['Story', 'Maximum'],
        #                                        filterKeyList=['OutputCase'],
        #                                        required_row=[['EQ']])


        U = ObtainResults.get_lateralDeflectionIndex(SapModel)

        # Q=ObtainResults.get_equivalentMomentofInertia(SapModel, storyName="Story1")


        # print(ans)
        # print(ans2)
        # retdata = ObtainResults.get_storyStiffness(SapModel)
        # print(sum([float(retdta) for retdta in retdata]))
        SapModel = None
        myETABSObject = None


# ======================================================================================================================


if not AttachToInstance:
    modelInfoAll, sectionInfoAll = Obj_utils.getModelSetting(filename="model.xlsx", sheetname='setting')

    resultData = np.zeros((np.shape(modelInfoAll)[0], 3))
    for i in range(np.shape(modelInfoAll)[0]):
        modelInfo = modelInfoAll[i, :]
        sectionInfo = sectionInfoAll[i, :]
        model1 = StructureModel()
        modelname = "MODEL_" + str(i + 1) + ".edb"
        resultData[i, :] = model1.createModel(modelInfo, sectionInfo, modelname)

    Obj_utils.writedata_excel(resultData, "results.xlsx", "result")
else:
    model1 = StructureModel()
    model1.runActiveModel()
