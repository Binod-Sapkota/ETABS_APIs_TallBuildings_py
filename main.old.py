import sys

import comtypes.client
# import comtypes.gen
import numpy as np
import os
import pandas as pd

import NewModel_utils
import Obj_utils
import OAPIs
import ObtainResults


# =========== Program Settings =========================================================================
resultfilename = "E:\\00_Thesis Project\ETABS_APIs_PYTHON\excelfiles\\results.xlsx"
AttachToInstance = False


SpecifyPath = False
ProgramPath = "C:\Program Files\Computers and Structures\ETABS 18\ETABS.exe"
APIPath = 'E:\\00_Thesis Project\ETABS MODELS'


# ======================================================================================================


def main():

    if AttachToInstance:
        try:
            # get the active ETABS object
            myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        except (OSError, comtypes.COMError):
            print("No running instance of the program found or failed to attach.")
            sys.exit(-1)

    else:
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
    ret = SapModel.SetPresentUnits(10)



    EtabsFile = Obj_utils.ObjEtabsFile("E:\\00_Thesis Project\Tall Building Models")
    etabsfiles = EtabsFile.files
    fileroot = EtabsFile.roots
    nameBuilding = EtabsFile.names


    nStory = []
    for f,rot,nam in zip(etabsfiles,fileroot,nameBuilding):
        ret = SapModel.File.OpenFile(f)
        numberStory,nameStory = OAPIs.get_NumberOfStory(SapModel)
        BRI = ObtainResults.getBRI(SapModel)
        dirs = os.path.basename(rot)
        savePath = os.path.join("E:\\00_Thesis Project\ETABS19Files", dirs)
        modelPath = savePath + os.sep + nam
        print(savePath)
        if not os.path.exists(savePath):
            try:
                os.makedirs(savePath)
            except OSError:
                pass
        ret = SapModel.File.Save(modelPath)

        COLUMNS = ['BRI']
        result = [BRI]

        result_df = pd.DataFrame(data=[result], index=None, columns=COLUMNS)

        # resultAll_df = resultAll_df.append(result_df, ignore_index=True)
        #
        # resultAll_df.to_excel(resultfilename, sheet_name="result")

    # print([nStory,nameBuilding])
    #
    # df = pd.DataFrame({'name': nameBuilding,
    #                    'numberOfStory':nStory})
    # df.to_csv("bldgInfo.csv", sep=",",index = False)
    #

    myETABSObject.ApplicationExit(False)
    SapModel = None
    myETABSObject = None

    return result_df


# ======================================================================================================================

main()
