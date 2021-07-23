import sys
import pandas as pd
import numpy as np
# from scipy.spatial import ConvexHull, convex_hull_plot_2d
import matplotlib.pyplot as plt
import comtypes.client
import ObtainResults
import NewModel_utils
# import test_Obj_utils
import switcher as sw
import alphashape
import switcher
import logging

logging.basicConfig(filename='test.log', level=logging.WARNING)


def runActiveModelforTest():
    try:
        # get the active ETABS object
        myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)

    # create SapModel object
    SapModel = myETABSObject.SapModel
    ret = SapModel.SetPresentUnits(10)

    # ============ Edit Below this ================================================

    # ObtainResults.obtainResults(SapModel, "test.xlsx", nrowGlobal=4, unitName="bldg",resType = "Global")

    # tableKey="Story Stiffness"
    # fields = SapModel.DatabaseTables.GetAllFieldsInTable(TableKey=tableKey)

    # switcher = sw.Switcher()
    # for i in range(21):
    #     result = switcher.func_to_call(i, SapModel)
    #     print(result)

    # test2 = ObtainResults.get_storyStiffnessIrregularity(SapModel)

    # [_, _, StoryName, storyelev, sHeight, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()
    #
    # storyNameList = []
    # for story in StoryName:
    #     try:
    #         [num, pname, ret] = SapModel.PointObj.GetNameListOnStory(story)
    #         [num, pname, ret] = SapModel.AreaObj.GetNameListOnStory(story)
    #         storyNameList.append(story)
    #     except:
    #         pass

    # ObtainResults.obtainResults(SapModel, "test.xlsx", nrowGlobal=2)

    # ret = ObtainResults.get_structuralPlanDensityIndex(SapModel)
    #
    # ret = ObtainResults.get_massStory(SapModel)

    # NewModel_utils.setLoadPtrnCaseCombosForOutput(SapModel, rmCase=(3, 10), rmPtrn=None, rmAllCombo=False)
    ObtainResults.allstories(SapModel)
    # ret = ObtainResults.get_planDimension(SapModel)

    #
    # ret = ObtainResults.get_InherentTorsionRatio(SapModel)

    # ret2 = ObtainResults.get_structuralPlanDensityIndex(SapModel)
    # ret3 = ObtainResults.get_storyStiffnessIrregularity(SapModel)

    # ret4 = ObtainResults.get_storyMassIrregularity(SapModel)

    # ret5 = ObtainResults.get_perimeterIndexes(SapModel)

    # ret6 = ObtainResults.get_vibrationSeriviceabilityIndex(SapModel)
    # ret7 = ObtainResults.get_massStory(SapModel)

    # ret2 = ObtainResults.get_structuralPlanDensityIndex(SapModel)

    ret3 = ObtainResults.get_perimeterIndexes(SapModel)
    return


runActiveModelforTest()
