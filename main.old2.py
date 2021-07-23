import sys
import os
import comtypes.client
import comtypes
import numpy as np

import pandas as pd
import json
from time import time
import NewModel_utils
import ObtainResults
import switcher as sw

# =========== Program Settings =========================================================================
datafilename = "D:\\00_Thesis Project\ETABS_APIs_PYTHON\modelInputfiles\model.xlsx"
resultfilename = "D:\\00_Thesis Project\ETABS_APIs_PYTHON\excelfiles\\results.xlsx"
AttachToInstance = True
runmode = "D"  # "I" - attach to instance, "D" - run from Directory, "N"- new model
nModels = 1

SpecifyPath = False
ProgramPath = "C:\Program Files\Computers and Structures\ETABS 18\ETABS.exe"
APIPath = 'E:\\00_Thesis Project\ETABS MODELS'

modelDirectory = "E:\\00_Thesis Project\ETABS19Files"
saveDirectory = "E:\\00_Thesis Project\ETABS19Files"


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
        topRightCornerJoint = NewModel_utils.drawNewModel(SapModel, ModelPath, datafilename,
                                                          modelInfo, sectionInfo)

        # ============== Analysis Setup ===============================
        NewModel_utils.analysisSetup(SapModel)
        ret = NewModel_utils.setModalCase(SapModel)

        # ============== Get Results ==================================================
        ASI = ObtainResults.get_axialCapacityIndex(SapModel)
        BF = ObtainResults.get_buckingFactor(SapModel)
        ISFX, ISFY = ObtainResults.get_instabilitySafetyFactor(SapModel)
        OSFX, OSFY = ObtainResults.get_overtuningSafetyFactor(SapModel)
        totalMass = ObtainResults.get_massStory(SapModel, "Total")
        dynamic_prop = ObtainResults.getDynamicProperty(SapModel)
        # BCSR = ObtainResults.get_beamColumnStiffnessRatio(SapModel)
        BCSR = 0
        LDIX, LDIY = ObtainResults.get_lateralDeflectionIndex(SapModel)
        [ux, uy, dftx, dfty, RDIx, RDIy] = ObtainResults.getJointResults(SapModel)
        SI = ObtainResults.get_shorteningIndex(SapModel)

        BRI = ObtainResults.getBRI(SapModel)
        period_sec = float(dynamic_prop['period_sec']['Period'].iloc[0])
        participation100 = dynamic_prop['participation100']

        myETABSObject.ApplicationExit(False)
        SapModel = None
        myETABSObject = None

        COLUMNS = ['ASI', 'BF', 'ISFX', 'ISFY', 'OSFX', 'OSFY', 'totalMass', 'period_sec', 'participation100',
                   'LDIX', 'LDIY', 'ux', 'uy', 'roofDriftX', 'roofDriftY', 'roofDriftIndexX',
                   'roofDriftIndexY', 'SI', 'BCSR', 'BRI']
        result = [ASI, BF, ISFX, ISFY, OSFX, OSFY, totalMass, period_sec, participation100,
                  LDIX, LDIY, ux, uy, dftx, dfty, RDIx, RDIy, SI, BCSR, BRI]

        result_df = pd.DataFrame(data=[result], index=None, columns=COLUMNS)

        return result_df

    def runActiveModel(self):
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

        # switcher = sw.Switcher()
        # [numberStory, _, ret] = SapModel.Story.GetNameList()
        # COLUMNS = np.array(["Story"] + list(range(1, numberStory))).reshape(-1, 1)
        # NewModel_utils.writedata_excel2(COLUMNS, "teststorydata.xlsx",
        #                                 sheetname="test1", startrange=(1, 1), to_clear=True)
        #
        # colNum = 2
        # for i in range(1):
        #     result = switcher.func_to_call(i, SapModel)
        #     if result is None: continue
        #     if not result['resStory'] is None:
        #         NewModel_utils.writedata_excel2(result['colName'], "teststorydata.xlsx",
        #                                         sheetname="test1", startrange=(1, colNum))
        #         NewModel_utils.writedata_excel2(result['resStory'], "teststorydata.xlsx",
        #                                         sheetname="test1", startrange=(2, colNum))
        #         colNum = colNum + np.shape(result['resStory'])[1]



    def runFromDirectory(self):
        EtabsFile = NewModel_utils.ObjEtabsFile(modelDirectory)
        etabsfiles = EtabsFile.files
        fileroot = EtabsFile.roots
        nameBuilding = EtabsFile.names

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

        # Set units
        kN_m_C = 6


        nrow = 1
        for f, rot, nam in zip(etabsfiles, fileroot, nameBuilding):
            nrow=nrow+1

            dirs = os.path.basename(rot)
            savePath = os.path.join(saveDirectory, dirs)
            modelPath = savePath + os.sep + nam
            resPath = saveDirectory + os.sep + "storyResults"
            resStoryfilename = resPath + os.sep + nam.split(".")[0] + ".xlsx"

            print(savePath)

            ret = SapModel.File.OpenFile(f)

            # input("Enter Any Key to Continue!!")  # stop program here (check ETABS file before processing)

            # ========= if needed to save =====================================================
            # if not os.path.exists(savePath):
            #     try:
            #         os.makedirs(savePath)
            #     except OSError:
            #         pass
            # ret = SapModel.File.Save(modelPath)
            # ==================================================================================
            if not os.path.exists(resPath):
                try:
                    os.makedirs(resPath)
                except OSError:
                    pass

            ObtainResults.obtainResults(SapModel, resStoryfilename, nrowGlobal=nrow, unitName=nam)


        myETABSObject.ApplicationExit(False)
        SapModel = None
        myETABSObject = None


# ======================================================================================================================


if (runmode == "N"):
    modelInfoAll, sectionInfoAll = NewModel_utils.getModelSetting(filename=datafilename, sheetname='setting',
                                                                  numberOfModels=nModels)

    resultData = np.zeros((np.shape(modelInfoAll)[0], 13))
    resultAll_df = pd.DataFrame()
    for i in range(np.shape(modelInfoAll)[0]):
        modelInfo = modelInfoAll[i, :]
        sectionInfo = sectionInfoAll[i, :]
        model1 = StructureModel()
        modelname = "MODEL_" + str(i + 1) + ".edb"
        result_df = model1.createModel(modelInfo, sectionInfo, modelname)
        resultAll_df = resultAll_df.append(result_df, ignore_index=True)

        resultAll_df.to_excel(resultfilename, sheet_name="result")

    # NewModel_utils.writedata_excel(resultData, resultfilename, "result", startrange=(1, 0))
elif runmode == "I":
    start_time = time()
    model1 = StructureModel()
    model1.runActiveModel()
    end_time = time()
    print(f'Total time taken is {end_time-start_time} sec')
elif runmode == "D":
    model1 = StructureModel()
    model1.runFromDirectory()

    # resultData.to_excel("result.xlsx", sheet_name="result")

    # NewModel_utils.writedata_excel(resultData, resultfilename, "result",startrange=(1,0))
