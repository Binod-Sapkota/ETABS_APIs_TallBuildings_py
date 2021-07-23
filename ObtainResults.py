import os
import numpy as np
import pandas as pd
# import matplotlib.pyplot as plt
# from scipy.spatial import ConvexHull, convex_hull_plot_2d
import alphashape
# from descartes import PolygonPatch
import switcher as sw
import NewModel_utils

storyTower = 0
optalpha=0.25
stories = []


# ======================================================================================================================
def obtainResults(SapModel, filename, nrowGlobal=2, unitName="bldg",resType="Full"):
    # resType = Full (story+global), Global and Story
    GFileName = "resGlobal.xlsx"
    GSheetName = "PEFs"

    ret = SapModel.SetPresentUnits(10)
    # ret = NewModel_utils.edit_diaphragm(SapModel)
    # ret = NewModel_utils.setbucklingLoadCase(SapModel)
    ret = NewModel_utils.setLoadCasetoRun(SapModel, notRunCase=[6])  # 6=wind
    # ret = SapModel.Analyze.SetSolverOption_2(2, 0, 0)
    # ret = SapModel.Analyze.RunAnalysis()
    # ret = NewModel_utils.setLoadPtrnCaseCombosForOutput(SapModel, rmCase=(3, 10), rmPtrn=None)  # 3=Modal, 10=Buckling

    switcher = sw.Switcher()
    [numberStory, storyNames, ret] = SapModel.Story.GetNameList()
    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    COLUMNS = np.array(["Story"] + list(storyNamelist) + ["Base"]).reshape(-1, 1)

    if resType != "Global":
        [app, wb] = NewModel_utils.writedata_excel2(COLUMNS, filename=filename,
                                                    sheetname="resStory", startrange=(1, 1), to_clear=True)

    colNum = 2
    ncolGlobal = 2

    [app, wb] = NewModel_utils.writedata_excel2(np.array(["Units"]), filename=GFileName,
                                                sheetname=GSheetName, startrange=(1, 1))
    [app, wb] = NewModel_utils.writedata_excel2(np.array([unitName]), filename=GFileName,
                                                sheetname=GSheetName, startrange=(nrowGlobal, 1))

    allstories(SapModel)     #//not required for [10] otherwise required//
    for i in [0,13]: #range(22):
        result = switcher.func_to_call(i, SapModel)
        if result is None: continue
        if not result['resStory'] is None:

            if resType != "Global":
                [app, wb] = NewModel_utils.writedata_excel2(result['colName'], filename=filename,
                                                            sheetname="resStory", startrange=(1, colNum))
                [app, wb] = NewModel_utils.writedata_excel2(result['resStory'], filename=filename,
                                                            sheetname="resStory", startrange=(2, colNum))

            colNum = colNum + np.shape(result['resStory'])[1]

        if nrowGlobal == 2:
            [app, wb] = NewModel_utils.writedata_excel2(result['colGlobal'], filename=GFileName,
                                                        sheetname=GSheetName, startrange=(1, ncolGlobal))

        if result['resGlobal'] is not None:
            [app, wb] = NewModel_utils.writedata_excel2(result['resGlobal'], filename=GFileName,
                                                        sheetname=GSheetName, startrange=(nrowGlobal, ncolGlobal))
            ncolGlobal = ncolGlobal + len(result['resGlobal'])

    app.quit()
    app = None
    switcher = None

# ======================================================================================================================
def savejointResultstotempCSV(SapModel):
    if os.path.exists(os.path.abspath("tempdisp.csv")):
        os.remove(os.path.abspath("tempdisp.csv"))

    [_, _, storyNamelist, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    ALLjointDisp_df = get_dataFromTable(SapModel, tableKey="Joint Displacements",
                                        fieldKeyList=['Story', 'Ux', 'Uy', 'Uz'],
                                        required_column=['Story', 'Ux', 'Uy', 'Uz'],
                                        filterKeyList=None,
                                        required_row=None)

    ALLjointDisp_df.to_csv("tempdisp.csv", index=False)
    ALLjointDisp_df = None


# ======================================================================================================================
def get_dataFromTable(SapModel, tableKey, fieldKeyList, required_column, filterKeyList=None, required_row=None):
    # required_row must be in the form of List of List such as [['a']] or [['a','b'],['1','2']]
    # the function returns pandas dataframe

    groupName = ""
    # SapModel.DatabaseTables.SetLoadCasesSelectedForDisplay(LoadCaseList=loadCase)

    # fields = SapModel.DatabaseTables.GetAllFieldsInTable(TableKey=tableKey)  # TODO ...use this to find fieldKeys

    tabledata = None
    [fieldkeyList, tableversion, fieldkeyincluded, numberrecords, tabledata,
     ret] = SapModel.DatabaseTables.GetTableForDisplayArray(TableKey=tableKey, FieldKeyList=fieldKeyList,
                                                            GroupName=groupName)
    try:
        data = np.asarray(tabledata).reshape((numberrecords, len(fieldkeyincluded)))
    except:
        print(f"Data Table \" {tableKey} \" might not have any data")
        return

    data_df = pd.DataFrame(data, columns=fieldkeyincluded)

    required_df = pd.DataFrame(columns=fieldkeyincluded)
    if filterKeyList is None:
        if required_column is None:
            required_df = data_df
        else:
            required_df = data_df[required_column]
    else:
        for filtercol in filterKeyList:
            required_df = pd.DataFrame(columns=fieldkeyincluded)
            for ro in required_row[filterKeyList.index(filtercol)]:
                required_df = required_df.append(data_df[data_df[filtercol] == ro], ignore_index=True)
            data_df = required_df

        if not required_column is None:
            required_df = required_df[required_column]

    return required_df


# ======================================================================================================================

class Column:
    def __init__(self, name, location, colSectionProp):
        self.name = name
        self.sectionProp = colSectionProp
        [self.ilocation, self.jlocation, self.col_length] = self.get_loc(location)
        # self.sectionProp = self.get_SectionProp(colSectionProp)
        # self.area = self.get_sectionArea()
        # self.area = colSectionProp[5]
        [self.matProp, self.tX, self.tY, self.rebarMatProp,
         self.propType, self.area, self.I33, self.I22] = self.sectionProp

        self.Ixx, self.Iyy = self.get_momentofInertia()

    # def get_SectionProp(self, colSectionProp):
    #     self.matProp = colSectionProp[0]
    #     self.tX = colSectionProp[1]
    #     self.tY = colSectionProp[2]
    #     self.rebarMatProp = colSectionProp[3]
    #     self.propType = colSectionProp[4]
    #     return (self.matProp, self.tX, self.tY, self.rebarMatProp, self.propType)

    def get_loc(self, location):
        self.xlocI = location[0][0]
        self.ylocI = location[0][1]
        self.zlocI = location[0][2]

        self.xlocJ = location[1][0]
        self.ylocJ = location[1][1]
        self.zlocJ = location[1][2]

        self.col_length = np.sqrt((self.xlocJ - self.xlocI) ** 2 + (self.ylocJ - self.ylocI) ** 2
                                  + (self.zlocJ - self.zlocI) ** 2)

        return ([[self.xlocI, self.ylocI, self.zlocI],
                 [self.xlocJ, self.ylocJ, self.zlocJ], self.col_length])

    # def get_sectionArea(self):
    #     if self.propType == 8:
    #         return self.sectionProp[1] * self.sectionProp[2]
    #     elif self.propType == 9:
    #         return np.pi * self.tX * self.tX / 4
    #     else:
    #         return self.sectionProp[5]

    def get_momentofInertia(self):
        if self.propType == 8:
            self.Ixx = self.tX * self.tY ** 3 / 12.0
            self.Iyy = self.tY * self.tX ** 3 / 12.0
            return self.Ixx, self.Iyy
        elif self.propType == 9:
            self.Ixx = np.pi * self.tX ** 4 / 64
            self.Iyy = np.pi * self.tX ** 4 / 64
            return self.Ixx, self.Iyy
        else:
            return self.I33, self.I22
            #


# ===============================================================================================================

class Wall:
    def __init__(self, name, location, wallSectionProp):
        self.name = name
        [self.ilocation, self.jlocation, self.wall_length, self.wall_height, self.theta] = self.get_loc(location)
        self.sectionProp = self.get_SectionProp(wallSectionProp)
        self.area = self.get_sectionArea()
        self.Ixx, self.Iyy = self.get_momentofInertia()

    def get_SectionProp(self, wallSectionProp):
        self.matProp = wallSectionProp[0]
        self.thickness = wallSectionProp[1]

        return (self.matProp, self.thickness)

    def get_loc(self, location):
        self.xlocI = location[0][0]
        self.ylocI = location[0][1]
        self.zlocI = location[0][2]

        self.xlocJ = location[1][0]
        self.ylocJ = location[1][1]
        self.zlocJ = location[1][2]

        self.xlocK = location[2][0]
        self.ylocK = location[2][1]
        self.zlocK = location[2][2]

        maxX = np.nanmax(location[:, 0])
        minX = np.nanmin(location[:, 0])
        maxY = np.nanmax(location[:, 1])
        minY = np.nanmin(location[:, 1])
        # maxZ = np.max(location[:][2])
        # minZ = np.min(location[:][2])
        #
        #
        self.sizeX = maxX - minX
        self.sizeY = maxY - minY
        # self.sizeZ = maxZ-minZ

        wall_length1 = np.sqrt((self.xlocJ - self.xlocI) ** 2 + (self.ylocJ - self.ylocI) ** 2
                               + (self.zlocJ - self.zlocI) ** 2)

        wall_length2 = np.sqrt((self.xlocK - self.xlocI) ** 2 + (self.ylocK - self.ylocI) ** 2)

        self.wall_height = np.sqrt((self.xlocK - self.xlocJ) ** 2 + (self.ylocK - self.ylocJ) ** 2
                                   + (self.zlocK - self.zlocJ) ** 2)

        self.wall_length = np.maximum(wall_length1, wall_length2)
        if self.sizeX > self.wall_length:
            self.wall_length = self.sizeX
        self.theta = 180.0 - np.degrees(np.arccos(self.sizeX / self.wall_length))

        self.ilocation = [(self.xlocI + self.xlocJ) / 2, (self.ylocI + self.ylocJ) / 2, (self.zlocI + self.zlocJ) / 2]
        self.jlocation = [(self.xlocI + self.xlocJ) / 2, (self.ylocI + self.ylocJ) / 2, self.zlocK]

        return ([self.ilocation,
                 self.jlocation, self.wall_length, self.wall_height, self.theta])

    def get_sectionArea(self):
        return self.sectionProp[1] * self.wall_length

    def get_momentofInertia(self):
        Ixx = self.sectionProp[1] * self.wall_length ** 3 / 12.0
        Iyy = self.wall_length * self.sectionProp[1] ** 3 / 12.0

        thetaYY = self.theta
        thetaXX = self.theta - 90.0

        self.Ixx = np.abs(Ixx * np.cos(np.radians(thetaXX)) + Iyy * np.sin(np.radians(thetaXX)))
        self.Iyy = np.abs(Ixx * np.cos(np.radians(thetaYY)) + Iyy * np.sin(np.radians(thetaYY)))

        return self.Ixx, self.Iyy


# ===============================================================================================================


class Floor:
    def __init__(self, name, location, floorSectionProp):
        self.name = name
        self.sectionProp = self.get_SectionProp(floorSectionProp)
        self.area = self.get_floorArea(location)

    def get_SectionProp(self, floorSectionProp):
        self.matProp = floorSectionProp[0]
        self.thickness = floorSectionProp[1]

        return (self.matProp, self.thickness)

    def get_floorArea(self, location):
        sumTotal = 0.0
        for i in range(np.size(location, axis=0)):
            if i == (np.size(location, axis=0) - 1):
                temp1 = location[i, 0] * location[0, 1] - location[0, 0] * location[i, 1]
                sumTotal += temp1
            else:
                temp1 = location[i, 0] * location[i + 1, 1] - location[i + 1, 0] * location[i, 1]
                sumTotal += temp1

        areaOfPolygon = abs(sumTotal / 2.0)
        return areaOfPolygon


# ======================================================================================================================

def allstories(SapModel):
    [numberStory, storyNames, ret] = SapModel.Story.GetNameList()
    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    global stories
    for i, story in enumerate(storyNamelist):
        stories.append(Storry(SapModel, story,medStr = storyNamelist[int(len(storyNamelist)/2)]))


# ===============================================================================================================

class Storry:
    def __init__(self, SapModel, name, medStr):
        self.name = name
        self.medStr = medStr
        self.get_Info(SapModel)

    def get_Info(self, SapModel):
        try:
            [_, pointNameinStorylist, ret] = SapModel.PointObj.GetNameListOnStory(self.name)
            X = []
            Y = []
            Z = []
            for pointName in pointNameinStorylist:
                [x, y, z, ret] = SapModel.PointObj.GetCoordCartesian(Name=pointName, CSys="Global")
                X.append(x)
                Y.append(y)
                Z.append(z)
            self.XYZ = np.transpose(np.array([X, Y, Z]))
            self.points = np.transpose(np.array([X, Y]))

            max_x = np.nanmax(self.XYZ[:, 0])
            max_y = np.nanmax(self.XYZ[:, 1])
            min_x = np.nanmin(self.XYZ[:, 0])
            min_y = np.nanmin(self.XYZ[:, 1])
            self.LenX = max_x - min_x
            self.LenY = max_y - min_y

            f = lambda x: alphashape.alphashape(self.points, x).area
            #
            global optalpha
            optalpha = self.optAlpha(f, a=0, b=0.7, step=0.05)

            self.alpha_shape = alphashape.alphashape(self.points, alpha=0.25)

            self.alphaArea = self.alpha_shape.area
            self.perimeter = self.alpha_shape.length
            self.centroid = self.alpha_shape.centroid

            # ======================================================================
            if self.name == self.medStr:

                filename = os.path.basename(SapModel.GetModelFilename())[:-4]

                import matplotlib.pyplot as plt
                from descartes import PolygonPatch
                fig1, ax1 = plt.subplots()
                ax1.scatter(*zip(*self.points))
                # ash = alphashape.alphashape(self.points)
                ax1.add_patch(PolygonPatch(self.alpha_shape, alpha=0.5))
                savename = "E:\\00_Thesis Project\\figures\\storyplanplots\\" + filename +".png"
                plt.savefig(savename,dpi=300,bbox_inches='tight')
                plt.close()
                # plt.show()
            # ======================================================================

            try:
                floorArea_df = get_dataFromTable(SapModel, tableKey="Material List by Story",
                                                 fieldKeyList=['Story', 'FloorArea'],
                                                 required_column=['Story', 'FloorArea'],
                                                 filterKeyList=['Story'],
                                                 required_row=[[self.name]])
                try:
                    self.tableArea = float(floorArea_df.drop_duplicates().iloc[0, 1])
                except:
                    self.tableArea = 0.0
            except:
                self.tableArea = 0.0

            if self.tableArea < self.alphaArea:
                self.Area = self.alphaArea
            else:
                self.Area = self.tableArea

        except:
            pass

    def optAlpha(self, f, a, b, step):
        a_n = []
        for n in np.arange(a, b + step, step):
            a_n.append(f(n))

        d = pd.Series(a_n)
        ma = d.rolling(3).mean()
        md = ma.rolling(3).apply(lambda x: x.iloc[0] - x.iloc[1])
        # print(d)
        # print(md)
        ind = np.nanargmax(md) - 1
        alpha = np.arange(a, b + step, step)[ind]
        # alpha = 0.1
        return alpha


# ===============================================================================================================
class Beam:
    def __init__(self, name, location, beamSectionProp):
        self.name = name
        [self.ilocation, self.jlocation, self.beam_length] = self.get_loc(location)
        self.sectionProp = self.get_SectionProp(beamSectionProp)
        self.area = self.get_sectionArea()
        self.I = self.get_momentofInertia()

    def get_SectionProp(self, beamSectionProp):
        self.matProp = beamSectionProp[0]
        self.depth = beamSectionProp[1]
        self.width = beamSectionProp[2]
        return (self.matProp, self.depth, self.width)

    def get_loc(self, location):
        self.xlocI = location[0][0]
        self.ylocI = location[0][1]
        self.zlocI = location[0][2]

        self.xlocJ = location[1][0]
        self.ylocJ = location[1][1]
        self.zlocJ = location[1][2]

        self.beam_length = np.sqrt((self.xlocJ - self.xlocI) ** 2 + (self.ylocJ - self.ylocI) ** 2
                                   + (self.zlocJ - self.zlocI) ** 2)

        return ([[self.xlocI, self.ylocI, self.zlocI],
                 [self.xlocJ, self.ylocJ, self.zlocJ], self.beam_length])

    def get_sectionArea(self):
        return self.sectionProp[1] * self.sectionProp[2]

    def get_momentofInertia(self):
        self.I = self.sectionProp[2] * self.sectionProp[1] ** 3 / 12.0
        return self.I


# ======================================================================================================================


def getStoryColumnInfo(SapModel, story="Story1"):
    colNames = []
    colLocations = []
    colSectionProp = []

    try:
        [_, frmNames, ret] = SapModel.FrameObj.GetNameListOnStory(story)
    except:
        return None

    for frm in frmNames:
        [designOrientation, ret] = SapModel.FrameObj.GetDesignOrientation(frm)
        if designOrientation == 1:
            colNames.append(frm)

    if not colNames:
        return None

    [_, frmPropName, frmPropType, t3, t2, _, _, _, _, frmPropArea, ret] = SapModel.PropFrame.GetAllFrameProperties_2()
    prop_df = pd.DataFrame(np.transpose(np.array([frmPropName, frmPropType, t3, t2, frmPropArea])),
                           columns=['frmPropName', 'frmPropType', 't3', 't2', 'frmPropArea'])

    columnNames = []
    for col in colNames:
        [colSection, _, ret] = SapModel.FrameObj.GetSection(col)

        [propType, ret] = SapModel.PropFrame.GetTypeOAPI(colSection)

        MatProp = None
        tX = 0
        tY = 0
        I33 = None
        I22 = None
        dff = prop_df[prop_df['frmPropName'] == colSection]
        propArea = float(dff.iloc[0, 4])
        if propType == 8:
            [_, MatProp, tX, tY, _, _, _, ret] = SapModel.PropFrame.GetRectangle(colSection)
        elif propType == 9:
            [_, MatProp, Dia, _, _, _, ret] = SapModel.PropFrame.GetCircle(colSection)
            tX = Dia
            tY = Dia
        else:
            tX = float(dff.iloc[0, 2])
            tY = float(dff.iloc[0, 3])
            [propArea, _, _, _, I22, I33, _, _, _, _, _, _, ret] = SapModel.PropFrame.GetSectProps(colSection)
            [MatProp, ret] = SapModel.PropFrame.GetMaterial(colSection)

            # print("Column section is other than rectangle or circle!! Please check")
            # continue

        columnNames.append(col)
        [rebarMatPropLong, _, _, _, _, _,
         _, _, _, _, _, _, _, _, ret] = SapModel.PropFrame.GetRebarColumn(colSection)

        if ret != 0:
            rebarMatPropLong = "A615Gr60"
        colSectionProp.append([MatProp, tX, tY, rebarMatPropLong, propType, propArea, I33, I22])

        [pointI, pointJ, ret] = SapModel.FrameObj.GetPoints(col)
        [xI, yI, zI, ret] = SapModel.PointObj.GetCoordCartesian(pointI)
        [xJ, yJ, zJ, ret] = SapModel.PointObj.GetCoordCartesian(pointJ)

        XYCO = [[xI, yI, zI], [xJ, yJ, zJ]]
        colLocations.append(XYCO)  # colLocation is 5 (column number)x2(i and j point)x3(x,y,z) list

    # column_dict = {"colNames":colNames,
    #                "colSectionProp":colSectionProp,
    #                "colLocations":colLocations}
    columns = []
    for i in range(len(columnNames)):
        columns.append(Column(columnNames[i], colLocations[i], colSectionProp[i]))
        # print(columns[i].col_length)
    return columns


# ======================================================================================================================

def getStoryWallInfo(SapModel, story='Story1'):
    wallNames = []
    wallLocations = []
    wallSectionProp = []

    try:
        [_, areaNames, ret] = SapModel.AreaObj.GetNameListOnStory(story)
    except:
        return None

    for ar in areaNames:
        [designOrientation, ret] = SapModel.AreaObj.GetDesignOrientation(ar)
        if designOrientation == 1:
            wallNames.append(ar)

    if not wallNames:
        return None

    for wl in wallNames:
        [wallProperty, ret] = SapModel.AreaObj.GetProperty(wl)
        [_, _, MatProp, thickness, _, _, _, ret] = SapModel.PropArea.GetWall(wallProperty)
        [npoints, points, ret] = SapModel.AreaObj.GetPoints(wl)
        x_co = []
        y_co = []
        z_co = []
        for pt in points:
            [x, y, z, ret] = SapModel.PointObj.GetCoordCartesian(pt)
            x_co.append(x)
            y_co.append(y)
            z_co.append(z)
        vertices = np.transpose(np.array([x_co, y_co, z_co]))
        wallSectionProp.append([MatProp, thickness])

        wallLocations.append(vertices)  # wallLocation is 5 (wall number)x4(vertices)x3(x,y,z) list

    walls = []
    for i in range(len(wallNames)):
        walls.append(Wall(wallNames[i], wallLocations[i], wallSectionProp[i]))
        # print(columns[i].col_length)
    return walls


# ======================================================================================================================

def getStoryfloorInfo(SapModel, story="Story1"):
    floorNames = []
    floorLocations = []
    floorSectionProp = []

    try:
        [_, areaNames, ret] = SapModel.AreaObj.GetNameListOnStory(story)
    except:
        return None

    for ar in areaNames:
        [designOrientation, ret] = SapModel.AreaObj.GetDesignOrientation(ar)
        if designOrientation == 2:
            floorNames.append(ar)

    if not floorNames:
        return None

    for wl in floorNames:
        [floorProperty, ret] = SapModel.AreaObj.GetProperty(wl)
        [_, _, MatProp, thickness, _, _, _, ret] = SapModel.PropArea.GetSlab(floorProperty)
        [npoints, points, ret] = SapModel.AreaObj.GetPoints(wl)
        x_co = []
        y_co = []
        z_co = []
        for pt in points:
            [x, y, z, ret] = SapModel.PointObj.GetCoordCartesian(pt)
            x_co.append(x)
            y_co.append(y)
            z_co.append(z)
        vertices = np.transpose(np.array([x_co, y_co, z_co]))
        floorSectionProp.append([MatProp, thickness])

        floorLocations.append(vertices)  # floorLocation is 5 (floor number)x4(vertices)x3(x,y,z) list

    floors = []
    for i in range(len(floorNames)):
        floors.append(Floor(floorNames[i], floorLocations[i], floorSectionProp[i]))
        # print(columns[i].col_length)

    floorArea = np.sum([ar.area for ar in floors])
    return floors, floorArea


# ======================================================================================================================

def getStoryBeamInfo(SapModel, story="Story1"):
    beamNames = []
    beamLocations = []
    beamSectionProp = []

    try:
        [_, frmNames, ret] = SapModel.FrameObj.GetNameListOnStory(story)
    except:
        return None

    for frm in frmNames:
        [designOrientation, ret] = SapModel.FrameObj.GetDesignOrientation(frm)
        if designOrientation == 2:
            beamNames.append(frm)
    if not beamNames:
        return None

    for beam in beamNames:
        [beamSection, _, ret] = SapModel.FrameObj.GetSection(beam)
        [_, MatProp, depth, width, _, _, _, ret] = SapModel.PropFrame.GetRectangle(beamSection)
        beamSectionProp.append([MatProp, depth, width])

        [pointI, pointJ, ret] = SapModel.FrameObj.GetPoints(beam)
        [xI, yI, zI, ret] = SapModel.PointObj.GetCoordCartesian(pointI)
        [xJ, yJ, zJ, ret] = SapModel.PointObj.GetCoordCartesian(pointJ)

        XYCO = [[xI, yI, zI], [xJ, yJ, zJ]]
        beamLocations.append(XYCO)  # colLocation is 5 (column number)x2(i and j point)x3(x,y,z) list

    beams = []
    for i in range(len(beamNames)):
        beams.append(Beam(beamNames[i], beamLocations[i], beamSectionProp[i]))
        # print(columns[i].col_length)
    return beams


# ======================================================================================================================


def getStoryInfo(SapModel, story):
    # [_, _, storyNamelist, elev, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()
    try:
        floorArea_df = get_dataFromTable(SapModel, tableKey="Material List by Story",
                                         fieldKeyList=['Story', 'FloorArea'],
                                         required_column=['Story', 'FloorArea'],
                                         filterKeyList=['Story'],
                                         required_row=[[story]])
        try:
            tableArea = float(floorArea_df.drop_duplicates().iloc[0, 1])
        except:
            tableArea = 0.0
            pass

        [_, pointNameinStorylist, ret] = SapModel.PointObj.GetNameListOnStory(story)
        X = []
        Y = []
        Z = []
        for pointName in pointNameinStorylist:
            [x, y, z, ret] = SapModel.PointObj.GetCoordCartesian(Name=pointName, CSys="Global")
            X.append(x)
            Y.append(y)
            Z.append(z)
        points = np.transpose(np.array([X, Y]))
        alpha_shape = alphashape.alphashape(points, 0.3)

        # fig, ax = plt.subplots()
        # ax.scatter(*zip(*points))
        # ax.add_patch(PolygonPatch(alpha_shape, alpha=0.5))
        # plt.show()

        Area = alpha_shape.area

        if Area < tableArea:
            return tableArea

    except:
        return None
    return Area


# ======================================================================================================================


def cantileverDeflection(EI, nStory, storyHeight, storyShear):
    # Unit used is N_m_C
    L = storyHeight
    node_force = storyShear
    # EI = np.dot(E, I)
    NELE = nStory
    EI = np.array(EI)

    a12s = 12 * np.divide(EI, np.power(L, 3))
    a6s = 6 * np.divide(EI, np.power(L, 2))
    a2s = 2 * np.divide(EI, L)
    a4s = 4 * np.divide(EI, L)

    element_stiffness = []

    for a12, a6, a2, a4 in zip(a12s, a6s, a2s, a4s):
        element_stiffness.append(np.array([[a12, a6, -a12, a6],
                                           [a6, a4, -a6, a2],
                                           [-a12, -a6, a12, -a6],
                                           [a6, a2, -a6, a4]]))

    # print(element_stiffness[0])

    sys_stiff = np.zeros((2 * (NELE + 1), 2 * (NELE + 1)))
    countEle = 0
    for elem_stiff in element_stiffness:
        for i in range(4):
            for j in range(4):
                sys_stiff[i + 2 * countEle, j + 2 * countEle] = sys_stiff[i + 2 * countEle, j + 2 * countEle] + \
                                                                elem_stiff[i, j]

        countEle += 1

    # print(sys_stiff)
    # sys_stiff_df = pd.DataFrame(sys_stiff)
    # sys_stiff_df.to_csv("sys_stiff1.csv")

    # Applying Boundary Condition using row column elimination

    sys_stiff = np.delete(sys_stiff, [0, 1], axis=0)
    sys_stiff = np.delete(sys_stiff, [0, 1], axis=1)
    node_force = np.delete(node_force, [0, 1], axis=0)

    U = np.matmul(np.linalg.inv(sys_stiff), node_force)
    # U is deflection of equivalent cantilever at story level in the direction of storyshear
    return U


# ======================================================================================================================
def get_equivalentBendingRigidity(SapModel, storyName="10F"):
    columns = getStoryColumnInfo(SapModel, story=storyName)

    walls = getStoryWallInfo(SapModel, story=storyName)

    if columns is None and walls is None:
        return

    if walls is None:
        l = len(columns)
    elif columns is None:
        l = len(walls)
    else:
        l = len(columns) + len(walls)

    matProp = []
    Ixx = np.zeros(l)
    Iyy = np.zeros(l)
    h = np.zeros(l)
    colArea = np.zeros(l)
    ilocation = np.zeros((l, 3))
    jlocation = np.zeros((l, 3))

    i = 0
    if (columns is not None) and (len(columns) != 0):
        for i, col in enumerate(columns):
            matProp.append(col.matProp)
            Ixx[i] = col.Ixx
            Iyy[i] = col.Iyy
            h[i] = col.col_length
            colArea[i] = col.area
            ilocation[i] = col.ilocation
            jlocation[i] = col.jlocation

    if not (walls is None):
        i = 0 if ((columns is None) or (len(columns) == 0)) else i + 1
        for wl in walls:
            matProp.append(wl.sectionProp[0])
            Ixx[i] = wl.Ixx
            Iyy[i] = wl.Iyy
            h[i] = wl.wall_height
            colArea[i] = wl.area
            ilocation[i] = wl.ilocation
            jlocation[i] = wl.jlocation
            i = i + 1

    E1 = get_dataFromTable(SapModel, tableKey="Material Properties - Basic Mechanical Properties",
                           fieldKeyList=['Material', 'E1'],
                           required_column=['E1'],
                           filterKeyList=['Material'],
                           required_row=[matProp])

    # E = E1['E1'].to_numpy(dtype=float)

    l = [i for i, e in enumerate(matProp) if e is None]
    E1 = E1['E1'].tolist()
    [E1.insert(i, 0) for i in l]

    E = np.asarray(E1, dtype=float)

    EI_hxx = np.sum(np.divide(np.multiply(E, Ixx), h))
    EI_hyy = np.sum(np.divide(np.multiply(E, Iyy), h))

    centroid_x = (np.nanmax(ilocation[:, 1]) - np.nanmin(ilocation[:, 1])) / 2
    centroid_y = (np.nanmax(ilocation[:, 0]) - np.nanmin(ilocation[:, 0])) / 2
    h_x = centroid_x - ilocation[:, 1]
    h_y = centroid_y - ilocation[:, 0]
    IXX = np.sum(Ixx + colArea * h_x ** 2)
    IYY = np.sum(Iyy + colArea * h_y ** 2)

    # Rigidity
    EIXX = np.sum(np.dot(E, Ixx + colArea * h_x ** 2))
    EIYY = np.sum(np.dot(E, Iyy + colArea * h_y ** 2))
    #        m^4 m^4   Nm^2  Nm^2   Nm      Nm
    return [IXX, IYY, EIXX, EIYY, EI_hxx, EI_hyy]


# ======================================================================================================================

def get_JointResults(SapModel):
    [_, _, storyNames, _, H, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    # U_df = get_dataFromTable(SapModel, tableKey="Diaphragm Center Of Mass Displacements",
    #                          fieldKeyList=['Story', 'OutputCase', 'UX', 'UY'],
    #                          required_column=['Story', 'UX', 'UY'],
    #                          filterKeyList=None,
    #                          required_row=None)

    # savejointResultstotempCSV(SapModel)

    U_df = get_dataFromTable(SapModel, tableKey="Joint Displacements - Absolute",
                             fieldKeyList=['Story', 'OutputCase', 'Ux', 'Uy'],
                             required_column=['Story', 'Ux', 'Uy'],
                             filterKeyList=None,
                             required_row=None)

    # U_df = pd.read_csv("tempdisp.csv")

    Ux = []
    Uy = []
    storyHeights = []
    storyNameList = []
    for i, story in enumerate(storyNamelist):
        df = U_df[U_df['Story'] == story]
        if not df.size == 0:
            storyHeights.append(H[i])
            storyNameList.append(storyNamelist[i])
            Ux.append(np.nanmax(abs(df["Ux"].to_numpy(dtype=float) * 1000)))  # converted to mm
            Uy.append(np.nanmax(abs(df["Uy"].to_numpy(dtype=float) * 1000)))  # converted to mm
    U_df = None
    HeightofStructure = sum(storyHeights)
    storyNumber = len(storyNameList)

    DriftX_df = get_dataFromTable(SapModel, tableKey="Story Drifts",
                                  fieldKeyList=['Story', 'Direction', 'Drift'],
                                  required_column=['Story', 'Direction', 'Drift'],
                                  filterKeyList=['Direction'],
                                  required_row=[['X']])

    DriftY_df = get_dataFromTable(SapModel, tableKey="Story Drifts",
                                  fieldKeyList=['Story', 'Direction', 'Drift'],
                                  required_column=['Story', 'Direction', 'Drift'],
                                  filterKeyList=['Direction'],
                                  required_row=[['Y']])

    DriftX = []
    DriftY = []
    for story in storyNameList:
        if not DriftX_df.empty:
            Xdf = DriftX_df[DriftX_df['Story'] == story]
            if not Xdf.empty:
                DriftX.append(np.nanmax(abs(Xdf["Drift"].to_numpy(dtype=float) * 100)))  # converted to %
            else:
                DriftX.append(0)
        if not DriftY_df.empty:
            Ydf = DriftY_df[DriftY_df['Story'] == story]
            if not Ydf.empty:
                DriftY.append(np.nanmax(abs(Ydf["Drift"].to_numpy(dtype=float) * 100)))  # converted to %
            else:
                DriftY.append(0)

    roofDriftX = 0
    roofDriftY = 0
    if not not DriftX:
        roofDriftX = DriftX[storyNumber - 1]
        if roofDriftX == 0.0:
            roofDriftX = DriftX[storyNumber - 2]

    if not not DriftY:
        roofDriftY = DriftY[storyNumber - 1]
        if roofDriftY == 0.0:
            roofDriftY = DriftY[storyNumber - 2]

    roofUx = Ux[storyNumber - 1]
    if roofUx == 0.0:
        roofUx = Ux[storyNumber - 2]

    roofUy = Uy[storyNumber - 1]
    if roofUy == 0.0:
        roofUy = Uy[storyNumber - 2]

    roofDriftIndexX = roofUx * 500 / HeightofStructure * (1 / 1000)  # Displacement converted back to meter
    roofDriftIndexY = roofUy * 500 / HeightofStructure * (1 / 1000)  # Displacement converted back to meter

    SDRX = np.nanmax(DriftX)
    SDRY = np.nanmax(DriftY)
    SDR = np.nanmax([SDRX, SDRY])

    # roof Drift Index should be less than one, the less the better(minimize)

    resDict = {"resStory": np.flip(np.transpose(np.array([Ux, Uy, DriftX, DriftY])), axis=0),
               'colName': ['Ux_mm', 'Uy_mm', 'DriftX_%', 'DriftY_%'],
               "resGlobal": [roofUx, roofUy, roofDriftX, roofDriftY, SDRX, SDRY, SDR, roofDriftIndexX, roofDriftIndexY],
               'colGlobal': ['roofUx', 'roofUy', 'roofDriftX_%', 'roofDriftY_%', 'SDRX', 'SDRY', 'SDR',
                             'roofDriftIndexX', 'roofDriftIndexY']}

    # return Ux,Uy,DriftX,DriftY,roofUx, roofUy, roofDriftX, roofDriftY, roofDriftIndexX, roofDriftIndexY
    return resDict


# from test import cantileverDeflection
# ======================================================================================================================

def get_lateralDeflectionIndex(SapModel):
    ret = SapModel.SetPresentUnits(10)
    [_, storyNumber, storyNames, _, H, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    IXX = []
    IYY = []
    EIXX = []
    EIYY = []
    storyHeights = []
    storyNameList = []
    for i, storyName in enumerate(storyNamelist):
        result = get_equivalentBendingRigidity(SapModel, storyName)
        if not (result is None):
            storyHeights.append(H[i])
            storyNameList.append(storyNamelist[i])
            Ixx, Iyy, EIxx, EIyy, _, _ = result
            IXX.append(Ixx)
            IYY.append(Iyy)
            EIXX.append(EIxx)
            EIYY.append(EIyy)

    storyShear_df = get_dataFromTable(SapModel, tableKey="Story Forces",
                                      fieldKeyList=['Story', 'Location', 'OutputCase', 'VX', 'VY', 'MX', 'MY'],
                                      required_column=['Story', 'VX', 'VY', 'MX', 'MY'],
                                      filterKeyList=['Location'],
                                      required_row=[['Bottom']])

    VX = []
    VY = []
    MX = []
    MY = []
    for story in storyNameList:
        df = storyShear_df[storyShear_df['Story'] == story]
        if not df.size == 0:
            VX.append(np.nanmax(abs(df["VX"].to_numpy(dtype=float))))  # converted to %
            VY.append(np.nanmax(abs(df["VY"].to_numpy(dtype=float))))  # converted to %
            MX.append(np.nanmax(abs(df["MX"].to_numpy(dtype=float))))
            MY.append(np.nanmax(abs(df["MY"].to_numpy(dtype=float))))
    # =========== x -direction ======================================================
    Data = np.flip(np.array(VX))
    storyNumber = len(VX)
    story_shear = np.zeros(len(Data))
    story_shear[0] = Data[0]
    for i in range(len(Data) - 1):
        story_shear[i + 1] = Data[i + 1] - Data[i]

    story_shear = np.flip(np.array(story_shear).astype(np.float))
    story_Shear = np.zeros((2 * (storyNumber + 1), 1))
    count = 0
    for i in range(2, 2 * (storyNumber + 1), 2):
        story_Shear[i, :] = story_shear[count]
        count += 1

    Ux = cantileverDeflection(EI=EIYY, nStory=storyNumber,
                              storyHeight=storyHeights, storyShear=story_Shear)

    lat_Cantileverdef = [Ux[2 * i][0] for i in range(storyNumber)]
    lat_CantileverdeflectionX = np.array(lat_Cantileverdef) * 1000

    # =========== y -direction ======================================================
    Data = np.flip(np.array(VY))

    story_shear = np.zeros(len(Data))
    story_shear[0] = Data[0]
    for i in range(len(Data) - 1):
        story_shear[i + 1] = Data[i + 1] - Data[i]

    story_shear = np.flip(np.array(story_shear).astype(np.float))
    story_Shear = np.zeros((2 * (storyNumber + 1), 1))
    count = 0
    for i in range(2, 2 * (storyNumber + 1), 2):
        story_Shear[i, :] = story_shear[count]
        count += 1

    Uy = cantileverDeflection(EI=EIXX, nStory=storyNumber,
                              storyHeight=storyHeights, storyShear=story_Shear)

    lat_Cantileverdef = [Uy[2 * i][0] for i in range(storyNumber)]
    lat_CantileverdeflectionY = np.array(lat_Cantileverdef) * 1000

    # Get maximum story Displacement
    # maxStoryDisplacementX_df = get_dataFromTable(SapModel, tableKey="Story Max Over Avg Displacements",
    #                                              fieldKeyList=['Story', 'OutputCase', 'Direction', 'Maximum'],
    #                                              required_column=['Story', 'Direction', 'Maximum'],
    #                                              filterKeyList=['Direction'],
    #                                              required_row=[['X']])

    df = get_dataFromTable(SapModel, tableKey="Joint Displacements - Absolute",
                           fieldKeyList=['Story', 'Ux', 'Uy'],
                           required_column=['Story', 'Ux', 'Uy'],
                           filterKeyList=None,
                           required_row=None)
    if df is None:
        df = get_dataFromTable(SapModel, tableKey="Joint Displacements",
                               fieldKeyList=['Story', 'Ux', 'Uy'],
                               required_column=['Story', 'Ux', 'Uy'],
                               filterKeyList=None,
                               required_row=None)

    # maxStoryDispX_df = df[df['Direction'] == "X"]
    # maxStoryDispY_df = df[df['Direction'] == "Y"]

    # maxStoryDisp_df = pd.read_csv("tempdisp.csv")

    disX = []
    disY = []

    for story in storyNameList:

        if df is not None:
            dfX = df[df['Story'] == story]
            if not dfX.empty:
                disX.append(np.nanmax(abs(dfX['Ux'].to_numpy(dtype=float) * 1000)))  # converted to mm
                disY.append(np.nanmax(abs(dfX['Uy'].to_numpy(dtype=float) * 1000)))  # converted to mm
        # if maxStoryDispY_df is not None:
        #     dfY = maxStoryDispY_df[maxStoryDispY_df['Story'] == story]
        #     if not dfY.empty:
        #         disY.append(max(abs(dfY['Maximum'].to_numpy(dtype=float) * 1000)))  # converted to mm

    # maxStoryDisp_df = None
    lateralDeflectionIndexX = 0
    lateralDeflectionIndexY = 0
    LDIX = []
    LDIY = []
    if not not disX:
        LDIX = np.divide(disX, lat_CantileverdeflectionX).reshape(-1, 1)
        lateralDeflectionIndexX = np.nanmean(LDIX)
    if not not disY:
        LDIY = np.divide(disY, lat_CantileverdeflectionY).reshape(-1, 1)
        lateralDeflectionIndexY = np.nanmean(LDIY)

    temp = np.transpose(np.concatenate((LDIX, LDIY), axis=1))

    resDict = {'resStory': np.flip(np.transpose(
        np.concatenate([np.divide([VX, VY, MX, MY], 1000), np.array([disX, disY]), temp], axis=0)), axis=0),
        'colName': ['storyShearX_KN', 'storyShearY_KN', 'storyMomentX_KNm', 'StoryMomentY_KNm',
                    'DisplacementX_mm', 'DisplacementY_mm', 'LDIX', 'LDIY'],
        'resGlobal': [lateralDeflectionIndexX, lateralDeflectionIndexY],
        'colGlobal': ['lateralDeflectionIndexX', 'lateralDeflectionIndexY']}

    return resDict


# ======================================================================================================================

def get_beamColumnStiffnessRatio(SapModel, storyName="Story1"):  # TODO check results and...generalize program
    # This factor might not be relevent in case of tall buildings with structural system other than moment frame

    beams = getStoryBeamInfo(SapModel, storyName)
    matProp = []
    depth = np.zeros((len(beams)))
    width = np.zeros((len(beams)))
    I = np.zeros((len(beams)))
    L = np.zeros((len(beams)))
    ilocation = np.zeros((len(beams), 3))
    jlocation = np.zeros((len(beams), 3))

    for i, beam in enumerate(beams):
        matProp.append(beam.sectionProp[0])
        depth[i] = beam.sectionProp[1]
        width[i] = beam.sectionProp[2]
        I[i] = beam.I
        L[i] = beam.beam_length
        ilocation[i] = beam.ilocation
        jlocation[i] = beam.jlocation

    E1 = get_dataFromTable(SapModel, tableKey="Material Properties - Basic Mechanical Properties",
                           fieldKeyList=['Material', 'E1'],
                           required_column=['E1'],
                           filterKeyList=['Material'],
                           required_row=[matProp])

    E = np.array(E1).astype(float)

    beam_area = depth * width
    AREA = np.sum(beam_area)

    EI_L = np.sum(np.divide(np.multiply(E[:, 0], I), L))
    [_, _, _, _, EI_hxx, EI_hyy] = get_equivalentBendingRigidity(SapModel, storyName=storyName)

    BCSR = EI_L / EI_hyy

    return BCSR


# ======================================================================================================================


def get_BRI(SapModel):
    [_, _, storyNames, elev, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    sBRI = []
    for _, story in enumerate(storyNamelist):  # i in range(1):   #

        columns = getStoryColumnInfo(SapModel, story=story)
        walls = getStoryWallInfo(SapModel, story=story)

        if columns is None and walls is None:
            continue

        if walls is None:
            l = len(columns)
        elif columns is None:
            l = len(walls)
        else:
            l = len(columns) + len(walls)

        Ixx = np.zeros(l)
        Iyy = np.zeros(l)
        h = np.zeros(l)
        colArea = np.zeros(l)
        ilocation = np.zeros((l, 3))
        jlocation = np.zeros((l, 3))

        i = 0
        if (columns is not None) and (len(columns) != 0):
            for i, col in enumerate(columns):
                Ixx[i] = col.Ixx
                Iyy[i] = col.Iyy
                h[i] = col.col_length
                colArea[i] = col.area
                ilocation[i] = col.ilocation
                jlocation[i] = col.jlocation

        if (walls is not None) and (len(walls) != 0):
            i = 0 if ((columns is None) or (len(columns) == 0)) else i + 1
            for wl in walls:
                Ixx[i] = wl.Ixx
                Iyy[i] = wl.Iyy
                h[i] = wl.wall_height
                colArea[i] = wl.area
                ilocation[i] = wl.ilocation
                jlocation[i] = wl.jlocation
                i = i + 1

        AREA = np.sum(colArea)
        centroid_x = (np.nanmax(ilocation[:, 1]) - np.nanmin(ilocation[:, 1])) / 2
        centroid_y = (np.nanmax(ilocation[:, 0]) - np.nanmin(ilocation[:, 0])) / 2
        h_x = centroid_x - ilocation[:, 1]
        h_y = centroid_y - ilocation[:, 0]
        IXX = np.sum(Ixx + colArea * h_x ** 2)
        IYY = np.sum(Iyy + colArea * h_y ** 2)

        if (not h_x.any()) or (not h_y.any()):
            n_corners = 2
        else:
            n_corners = 4

        A_eq = AREA / n_corners * np.ones(n_corners)
        d_eq = np.sqrt(A_eq) * np.ones(n_corners)
        b_eq = d_eq * np.ones(n_corners)

        corner_location = 0
        if n_corners == 2:
            corner_location = np.array([[np.nanmin(ilocation[:, 0]), np.nanmin(ilocation[:, 1])],
                                        [np.nanmax(ilocation[:, 0]), np.nanmin(ilocation[:, 1])]])
        elif n_corners == 4:
            corner_location = np.array([[np.nanmin(ilocation[:, 0]), np.nanmin(ilocation[:, 1])],
                                        [np.nanmax(ilocation[:, 0]), np.nanmin(ilocation[:, 1])],
                                        [np.nanmax(ilocation[:, 0]), np.nanmax(ilocation[:, 1])],
                                        [np.nanmin(ilocation[:, 0]), np.nanmax(ilocation[:, 1])]])

        I_eq = b_eq * d_eq ** 3 / 12

        if not h_x.any():
            h_yeq = centroid_y - corner_location[:, 0]
            IYY_eq = np.sum(I_eq + A_eq * h_yeq ** 2)
            BRI = IYY / IYY_eq
        elif not h_y.any():
            h_xeq = centroid_x - corner_location[:, 1]
            IXX_eq = np.sum(I_eq + A_eq * h_xeq ** 2)
            BRI = IXX / IXX_eq
        else:
            h_xeq = centroid_x - corner_location[:, 1]
            h_yeq = centroid_y - corner_location[:, 0]
            IXX_eq = np.sum(I_eq + A_eq * h_xeq ** 2)
            IYY_eq = np.sum(I_eq + A_eq * h_yeq ** 2)
            BRIx = IXX / IXX_eq
            BRIy = IYY / IYY_eq
            BRI = np.nanmin([BRIy, BRIx])
        sBRI.append(BRI)
    BRI = np.nanmean(sBRI)

    resDict = {'resStory': np.flip(sBRI).reshape(-1, 1),
               'colName': ['BendingRigidityIndex(BRI)'],
               'resGlobal': [BRI],
               'colGlobal': ['BRI']}

    return resDict


# ======================================================================================================================


def get_DynamicProperty(SapModel):
    # period_sec = get_dataFromTable(SapModel, tableKey="Modal Periods And Frequencies",
    #                                fieldKeyList=['Mode', 'Period'],
    #                                required_column=['Period'],
    #                                filterKeyList=['Mode'],
    #                                required_row=[['1', '2', '3']])

    df = get_dataFromTable(SapModel, tableKey="Modal Participating Mass Ratios",
                           fieldKeyList=['Mode', "Period", "UX", "UY", "RZ", 'SumUX', 'SumUY', 'SumRZ'],
                           required_column=['Mode', "Period", "UX", "UY", "RZ", 'SumUX', 'SumUY', 'SumRZ'],
                           filterKeyList=None,
                           required_row=None)
    T = []
    Tx = 0
    Ty = 0
    Tz = 0
    for i in range(1, 4):
        m = df.loc[df['Mode'] == str(i)]
        T.append(np.nanmax(m['Period'].to_numpy(dtype=float)))
        mX = np.nanmax(m['UX'].to_numpy(dtype=float))
        mY = np.nanmax(m['UY'].to_numpy(dtype=float))
        mZ = np.nanmax(m['RZ'].to_numpy(dtype=float))

        if mY < mX > mZ:
            Tx = m['Period'].to_numpy(dtype=float)
        elif mX < mY > mZ:
            Ty = m['Period'].to_numpy(dtype=float)
        elif mX < mZ > mY:
            Tz = m['Period'].to_numpy(dtype=float)

    m1 = df.loc[0:2]
    m1 = m1.to_numpy(dtype=float)
    maxInd = np.argmax(m1, axis=0)
    xind = maxInd[2]
    yind = maxInd[3]
    zind = maxInd[4]

    Tx1 = m1[xind, 1]
    Ty1 = m1[yind, 1]
    Tz1 = m1[zind, 1]

    TPRX = Tx1 / Tz1
    TPRY = Ty1 / Tz1

    # TPRX = Tx / Tz
    # TPRY = Ty / Tz

    TPR = np.nanmax([TPRX, TPRY])
    if (TPR - 1) > 0.05:
        Type = "Torsionally-stiff (TPR>1)"
    elif (1 - TPR) > 0.05:
        Type = "Torsionally-flexible (TPR<1)"
    else:
        Type = "Torsionally-similarly-stiff (TPR~1)"

    massParticipation = df[['SumUX', 'SumUY', "SumRZ"]].to_numpy(dtype=float)
    try:
        numberofModes = np.where(massParticipation > 0.999)[0][0]  # TODO ....change limit
    except:
        # import NewModel_utils
        # NewModel_utils.setModalCase(SapModel, nMode=50)
        try:
            numberofModes = np.where(massParticipation > 0.99)[0][0]
        except:
            numberofModes = massParticipation[-1, 1]

    resDict = {'resStory': None,
               'resGlobal': [T[0].item(), np.nanmax(massParticipation[0, :]), numberofModes, TPR.item(), Type],
               'colGlobal': ['T_sec', 'massParticipation', 'numberofModes', 'TPR', 'Type']}

    return resDict


# ======================================================================================================================

def get_planDimension(SapModel):
    # [StoryNumber, StoryName, ret] = SapModel.Story.GetNameList()
    [baseElevation, _, storyNames, elev_, heights_, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    elev = []
    heights = []
    for i, story in enumerate(storyNames):
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
            elev.append(elev_[i])
            heights.append(heights_[i])
        except:
            pass

    LenX = []
    LenY = []
    Elevations = []
    perimeter = []
    Area = []
    Area2 = []
    h = []

    floorArea_df = get_dataFromTable(SapModel, tableKey="Material List by Story",
                                     fieldKeyList=['Story', 'FloorArea'],
                                     required_column=['Story', 'FloorArea'],
                                     filterKeyList=['Story'],
                                     required_row=[storyNamelist])
    floorArea_df = floorArea_df.drop_duplicates()

    X1 = []
    Y1 = []
    Z1 = []
    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        # story = "44F"
        try:
            [_, pointNameinStorylist, ret] = SapModel.PointObj.GetNameListOnStory(story)
            X = []
            Y = []
            Z = []
            Elevations.append(elev[i])
            h.append(heights[i])
            for pointName in pointNameinStorylist:
                [x, y, z, ret] = SapModel.PointObj.GetCoordCartesian(Name=pointName, CSys="Global")
                X.append(x)
                Y.append(y)
                Z.append(z)
                X1.append(x)
                Y1.append(y)
                Z1.append(z)

            XYZ = np.transpose(np.array([X, Y, Z]))

            max_x = np.nanmax(XYZ[:, 0])
            max_y = np.nanmax(XYZ[:, 1])
            min_x = np.nanmin(XYZ[:, 0])
            min_y = np.nanmin(XYZ[:, 1])

            LenX.append(max_x - min_x)
            LenY.append(max_y - min_y)

            points = np.transpose(np.array([X, Y]))

            alpha_shape = alphashape.alphashape(points, alpha=optalpha)
            # alpha_shape = alphashape.alphashape(points)
            # =========================================================
            # if i == 1:
            #     fig, ax = plt.subplots()
            #     ax.scatter(*zip(*points))
            #     ax.add_patch(PolygonPatch(alpha_shape, alpha=0.5))
            #     # plt.show()

            # =========================================================

            perimeter.append(alpha_shape.length)
            Area.append(alpha_shape.area)

            # ========= Using Convex Hull ====================
            # hull = ConvexHull(points)
            # _ = convex_hull_plot_2d(hull)
            # plt.show()
            # perimeter.append(hull.area)
            # Area.append(hull.volume)
            # ===========================================
            if not floorArea_df.empty and not floorArea_df[floorArea_df['Story'] == story]['FloorArea'].empty:

                Area2.append((floorArea_df[floorArea_df['Story'] == story]['FloorArea']).to_numpy(dtype=float).item())
            else:
                Area2.append(0.0)
        except:
            perimeter.append(0.0)
            Area.append(0.0)
            Area2.append(0.0)

    points_3d = np.transpose(np.array([X1, Y1, Z1]))
    alpha_shape3d = alphashape.alphashape(points_3d, 0.3)
    # alpha_shape3d = alphashape.alphashape(points_3d)

    Volumn = alpha_shape3d.volume
    surfaceArea = alpha_shape3d.area
    SAVR = surfaceArea / Volumn

    # fig = plt.figure()
    # ax = plt.axes(projection='3d')
    # ax.plot_trisurf(*zip(*alpha_shape.vertices), triangles=alpha_shape.faces)
    # plt.show()

    # plt.show()
    # input("Press any key to continue")
    medLenX = np.median(LenX)
    medLenY = np.median(LenY)
    medperimeter = np.median(perimeter)
    medArea = np.median(Area)
    medArea2 = 0

    if not floorArea_df.empty:
        medArea2 = np.median(Area2)
        TFA = np.nansum(Area2)
        USAR = TFA / surfaceArea
    else:
        TFA = np.nansum(Area)
        USAR = TFA / surfaceArea

    nStories = len(h)
    totalHeight = np.nanmax(Elevations) - baseElevation
    buildingAspectRatio = totalHeight / np.minimum(np.median(LenX), np.median(LenY))
    # formatAspectRatio = f"1:{round(buildingAspectRatio,2)}"

    data = pd.DataFrame((storyNamelist, LenX, LenY, perimeter, Area, Area2)).T
    data.columns = ['Story', 'LenX', 'LenY', 'perimeter', 'Area', 'Area2']
    df = data.iloc[:, 1:6].copy()
    from sklearn.cluster import KMeans
    kmeans = KMeans(n_clusters=3, random_state=0)
    df['cluster'] = kmeans.fit_predict(df)
    df['Story'] = data['Story']

    temp = np.flip(df['cluster'].to_numpy())
    temp2 = np.argmax((np.sum(temp == 0), np.sum(temp == 1), np.sum(temp == 2)))
    temp3 = np.where(temp == temp2)[0][0]
    ind = len(temp) - temp3

    # podNum = temp3
    # podium_level = df['Story'].iloc[ind - 1]

    groups = df.groupby("cluster")
    grp = [g for n, g in groups]
    global storyTower
    storyTower = grp[temp2]['Story'].to_list()

    resDict = {'resStory': np.flip(np.transpose(np.array([h, LenX, LenY, perimeter, Area, Area2])), axis=0),
               'colName': ['StoryHeight_m', 'LengthX_m', 'LengthY_m', 'Perimeter_m', 'Area(from alpha shape)_m^2',
                           'Area(from table)_m^2'],
               'resGlobal': [nStories, totalHeight, medLenX, medLenY, medperimeter, medArea, medArea2, TFA,
                             buildingAspectRatio, Volumn, surfaceArea, SAVR, USAR],
               'colGlobal': ['nStories', 'Height_m', 'medLenX_m', 'medLenY_m', 'medperimeter_m', 'medArea(alpha)_m^2',
                             'medArea2(table)_m^2', 'TFA', 'buildingAspectRatio', 'Volumn_m^3',
                             'SurfaceArea', 'SAVR', 'USAR']}

    return resDict


# ======================================================================================================================


def get_massStory(SapModel):
    [_, storyNames, ret] = SapModel.Story.GetNameList()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    mass_kg = get_dataFromTable(SapModel, tableKey="Mass Summary by Group",
                                fieldKeyList=['Group', 'SelfWeight', 'MassX'],
                                required_column=['SelfWeight', 'MassX'],
                                filterKeyList=None,
                                required_row=None)

    totalMass_kg = np.nanmax(mass_kg['MassX'].to_numpy(dtype=float))
    elementMass_kg = np.nanmax(mass_kg['SelfWeight'].to_numpy(dtype=float))/9.81

    mass_kg = get_dataFromTable(SapModel, tableKey="Mass Summary by Story",
                                fieldKeyList=['Story', 'UX'],
                                required_column=['Story', 'UX'],
                                filterKeyList=['Story'],
                                required_row=[storyNamelist])
    storyMass_kg = mass_kg['UX'].to_numpy(dtype=float).reshape(-1, 1)

    # mass_kg = get_dataFromTable(SapModel, tableKey="Mass Summary by Diaphragm",
    #                             fieldKeyList=['Story', 'MassX'],
    #                             required_column=['MassX'],
    #                             filterKeyList=['Story'],
    #                             required_row=[storyNamelist])
    #
    # diaphragmMass_kg = mass_kg['MassX'].to_numpy(dtype=float).reshape(-1, 1)

    # return totalMass_kg, storyMass_kg, diaphragmMass_kg

    resDict = {'resStory': np.array([storyMass_kg]).reshape(-1, 1),
               'colName': ["storyMass_kg"],
               'resGlobal': [elementMass_kg, totalMass_kg, np.nanmean(storyMass_kg)],
               'colGlobal': ['elementMass_kg', 'total_mass_kg', 'averageMass_kg']}

    return resDict


# ======================================================================================================================

def get_overtuningSafetyFactor(SapModel):
    # return overturning moment due to selfweight and base shear

    [_, storyNames, ret] = SapModel.Story.GetNameList()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    # overturningMoment_kNm_df
    OM_df = get_dataFromTable(SapModel, tableKey="Story Forces",
                              fieldKeyList=['Story', 'OutputCase', 'MX', 'MY'],
                              required_column=['Story', 'OutputCase', 'MX', 'MY'],
                              filterKeyList=None,
                              required_row=None)

    # overturningMomentX_kNm = max(abs(overturningMoment_kNm_df['MX'].to_numpy(dtype=float)))
    # overturningMomentY_kNm = max(abs(overturningMoment_kNm_df['MY'].to_numpy(dtype=float)))

    # RF_df = get_dataFromTable(SapModel, tableKey="Story Forces",
    #                           fieldKeyList=['Story', 'OutputCase', 'P'],
    #                           required_column=['Story', 'P'],
    #                           filterKeyList=None,
    #                           required_row=None)

    # resistingForce_kNm = max(abs(resistingForce_kNm_df['P'].to_numpy(dtype=float)))

    CM_df = get_dataFromTable(SapModel, tableKey="Centers Of Mass And Rigidity",
                              fieldKeyList=['Story', 'Diaphragm', 'MassX', 'CumMassX', 'XCM', 'YCM'],
                              required_column=['Story', 'Diaphragm', 'MassX', 'CumMassX', 'XCM', 'YCM'],
                              filterKeyList=None,
                              required_row=None)

    OMX = []
    OMY = []
    RF = []
    CMX = []
    CMY = []

    for i, story in enumerate(storyNamelist):
        odf = OM_df[OM_df['Story'] == story]
        # rdf = RF_df[RF_df['Story'] == story]
        c_df = CM_df[CM_df['Story'] == story]
        if not odf.size == 0:
            # storyHeights.append(H[i])
            # storyNameList.append(StoryName[i])
            OMX.append(np.nanmax(abs(odf["MX"].to_numpy(dtype=float))))
            OMY.append(np.nanmax(abs(odf["MY"].to_numpy(dtype=float))))

            if not c_df.empty and sum(c_df['MassX'].to_numpy(dtype=float)) != 0:
                RF.append(np.nanmax(abs(c_df["CumMassX"].to_numpy(dtype=float))) * 9.81)
                CMX.append(abs(np.dot(c_df['MassX'].to_numpy(dtype=float), c_df['XCM'].to_numpy(dtype=float)) / sum(
                    c_df['MassX'].to_numpy(dtype=float))))

                CMY.append(abs(np.dot(c_df['MassX'].to_numpy(dtype=float), c_df['YCM'].to_numpy(dtype=float)) / sum(
                    c_df['MassX'].to_numpy(dtype=float))))
            else:
                RF.append(0.0)
                CMX.append(0.0)
                CMY.append(0.0)

    # XCCM = centerOfMass_m_df['XCM'].to_numpy(dtype=float)
    # YCCM = centerOfMass_m_df['YCM'].to_numpy(dtype=float)

    resistingMomentY_kNm = np.multiply(RF, CMX)
    resistingMomentX_kNm = np.multiply(RF, CMY)

    OSFX = 0
    OSFY = 0

    if np.any(resistingMomentY_kNm) and np.any(OMY):
        OSFX = np.divide(resistingMomentY_kNm, OMY)
    if np.any(resistingMomentX_kNm) and np.any(OMX):
        OSFY = np.divide(resistingMomentX_kNm, OMX)

    OSFX_ = 0.0
    OSFY_ = 0.0
    if np.any(resistingMomentY_kNm) and np.any(OMY):
        OSFX_ = np.divide(np.sum(resistingMomentY_kNm), np.sum(OMY))
    if np.any(resistingMomentX_kNm) and np.any(OMX):
        OSFY_ = np.divide(np.sum(resistingMomentX_kNm), np.sum(OMX))

    resDict = {'resStory': np.transpose([OSFX, OSFY]),
               'colName': ["OverturningSafetyFactorX", "OverturningSafetyFactorY"],
               'resGlobal': [OSFX_, OSFY_],
               'colGlobal': ['OSFX', 'OSFY']}

    return resDict


# ======================================================================================================================

def get_buckingFactor(SapModel):
    # return buckling factor for first mode
    # Buckling Factor is the ratio of axial load to euler critical buckling load
    # It is manually verified. Also for reference https://www.youtube.com/watch?v=sxyBcqhmQyI

    res = get_dataFromTable(SapModel, tableKey="Buckling Factors",
                            fieldKeyList=['Case', 'Mode', 'ScaleFactor'],
                            required_column=['ScaleFactor'],
                            filterKeyList=['Mode'],
                            required_row=[['1']])

    bucklingFactor = res['ScaleFactor'].to_numpy(dtype=float)

    # return bucklingFactor[0]
    resDict = {'resStory': None,
               'resGlobal': [bucklingFactor[0]],
               'colGlobal': ['bucklingFactor']}

    return resDict


# ======================================================================================================================

def get_modulusOfElasticity(SapModel):
    [_, matName, ret] = SapModel.PropMaterial.GetNameList()
    E = []
    fc = []
    for mat in matName:
        [matType, symType, ret] = SapModel.PropMaterial.GetTypeOAPI(mat)
        if matType == 2:
            [e, u, a, g, ret] = SapModel.PropMaterial.GetMPIsotropic(mat)
            E.append(e * 10 ** (-6))
            f = SapModel.PropMaterial.GetOConcrete(mat)[0]
            fc.append(f)
    maxE = np.nanmax(E)  # in MPa

    fck = fc[np.nanargmax(E)] * 10 ** (-6)

    # fck = (maxE / 4700) ** 2

    resDict = {'resStory': None,
               'resGlobal': [fck, maxE],
               'colGlobal': ['fck2', 'concE_MPa']}

    return resDict


# ======================================================================================================================

def get_baseShear(SapModel):
    # return base shear
    coeff_X = 0
    coeff_Y = 0
    baseShearES_X = 0
    baseShearES_Y = 0
    try:
        df = get_dataFromTable(SapModel, tableKey="Load Pattern Definitions",
                               fieldKeyList=['Type', "AutoLoad"],
                               required_column=['Type', "AutoLoad"],
                               filterKeyList=None,
                               required_row=None)

        # if df[df['Type'] == "Seismic"].empty:
        #     return None

        autoLoad = df[(df['Type'] == "Seismic") & (df['AutoLoad'] != "User Loads")]['AutoLoad']
        # if autoLoad.empty:
        #     return None
        autoLoad = autoLoad.iloc[0]
        seismicTable = "Load Pattern Definitions - Auto Seismic - " + autoLoad

        reqCol = ['XDir', 'XDirPlusE', 'XDirMinusE', 'YDir', 'YDirPlusE', 'YDirMinusE',
                  'CoeffUsed', 'WeightUsed', 'BaseShear']
        df = get_dataFromTable(SapModel, tableKey=seismicTable,
                               fieldKeyList=reqCol,
                               required_column=reqCol,
                               filterKeyList=None,
                               required_row=None)

        Xdf = df[(df['XDir'] == 'Yes') | (df['XDirPlusE'] == 'Yes') | (df['XDirMinusE'] == 'Yes')]
        Ydf = df[(df['YDir'] == 'Yes') | (df['YDirPlusE'] == 'Yes') | (df['YDirMinusE'] == 'Yes')]

        coeff_X = 0
        baseShearES_X = 0
        coeff_Y = 0
        baseShearES_Y = 0

        if not Xdf.empty:
            coeff_X = np.nanmax(Xdf['CoeffUsed'].to_numpy(dtype=float))
            baseShearES_X = np.nanmax(Xdf['BaseShear'].to_numpy(dtype=float)) / 1000
        if not Ydf.empty:
            coeff_Y = np.nanmax(Ydf['CoeffUsed'].to_numpy(dtype=float))
            baseShearES_Y = np.nanmax(Ydf['BaseShear'].to_numpy(dtype=float)) / 1000
    except:
        pass

    df = get_dataFromTable(SapModel, tableKey="Base Reactions",
                           fieldKeyList=['OutputCase', 'CaseType', 'FX', 'FY', 'MX', 'MY'],
                           required_column=['OutputCase', 'CaseType', 'FX', 'FY', 'MX', 'MY'],
                           filterKeyList=['CaseType'],
                           required_row=[["LinRespSpec"]])

    baseShearX = np.nanmax(df['FX'].to_numpy(dtype=float)) / 1000
    baseShearY = np.nanmax(df['FY'].to_numpy(dtype=float)) / 1000
    baseMomentX = np.nanmax(df['MX'].to_numpy(dtype=float)) / 1000
    baseMomentY = np.nanmax(df['MY'].to_numpy(dtype=float)) / 1000

    mass_kg = get_dataFromTable(SapModel, tableKey="Mass Summary by Group",
                                fieldKeyList=['Group', 'MassX'],
                                required_column=['MassX'],
                                filterKeyList=None,
                                required_row=None)

    totalMass_kg = np.nanmax(mass_kg['MassX'].to_numpy(dtype=float))

    baseShearXp = baseShearX * 1000 / (totalMass_kg * 9.81) * 100
    baseShearYp = baseShearY * 1000 / (totalMass_kg * 9.81) * 100

    resDict = {'resStory': None,
               'resGlobal': [coeff_X, coeff_Y, baseShearES_X, baseShearES_Y, baseShearX, baseShearY,
                             baseMomentX, baseMomentY, baseShearXp, baseShearYp],
               'colGlobal': ['coeff_X', 'coeff_Y', 'baseShearES_X_KN', 'baseShearES_Y_KN', 'baseShearX_KN',
                             'baseShearY_KN',
                             'baseMomentX_KNm', 'baseMomentY_KNm', 'baseShearX_%', 'baseShearY_%']}

    return resDict


# ======================================================================================================================

def get_instabilitySafetyFactor(SapModel):
    # return instability safety factor

    [baseElev, _, storyNames, storyelev_, sHeight_, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    storyelev = []
    sHeight = []
    for i, story in enumerate(storyNames):
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
            storyelev.append(storyelev_[i])
            sHeight.append(sHeight_[i])
        except:
            pass

    [_, caseNames, ret] = SapModel.LoadCases.GetNameList()
    # [_, patternNames, ret] = SapModel.LoadPatterns.GetNameList()

    Vx_df = get_dataFromTable(SapModel, tableKey="Story Forces",
                              fieldKeyList=['Story', 'Location', 'OutputCase', 'VX', 'VY'],
                              required_column=['Story', 'OutputCase', 'VX', 'VY'],
                              filterKeyList=['Location'],
                              required_row=[['Bottom']])

    # P_df = get_dataFromTable(SapModel, tableKey="Story Forces",
    #                          fieldKeyList=['Story', 'Location', 'OutputCase', 'P'],
    #                          required_column=['Story', 'P'],
    #                          filterKeyList=['Location'],
    #                          required_row=[['Bottom']])

    P_df = get_dataFromTable(SapModel, tableKey="Centers Of Mass And Rigidity",
                             fieldKeyList=['Story', 'Diaphragm', 'CumMassX', 'XCM', 'YCM'],
                             required_column=['Story', 'Diaphragm', 'CumMassX', 'XCM', 'YCM'],
                             filterKeyList=None,
                             required_row=None)

    driftX_df = get_dataFromTable(SapModel, tableKey="Diaphragm Center Of Mass Displacements",
                                  fieldKeyList=['Story', 'OutputCase', 'UX', 'UY'],
                                  required_column=['Story', 'OutputCase', 'UX', 'UY'],
                                  filterKeyList=None,
                                  required_row=None)

    Vx = []
    Vy = []
    h = []
    P = []
    driftX = []
    driftY = []
    storyIndex = []
    sh =[]

    for i, story in enumerate(storyNamelist):
        df1 = Vx_df[Vx_df['Story'] == story]
        df2 = P_df[P_df['Story'] == story]
        df3 = driftX_df[driftX_df['Story'] == story]

        if (df1.size != 0) and (df2.size != 0) and (df3.size != 0) \
                and max(abs(df2["CumMassX"].to_numpy(dtype=float))) != 0.0:
            storyIndex.append(i)
            h.append((storyelev[i] - baseElev) * 1000)  # mm
            sh.append(sHeight[i]*1000)
            Vx.append(np.nanmax(abs(df1["VX"].to_numpy(dtype=float) / 1000)))  # converted to KN
            Vy.append(np.nanmax(abs(df1["VY"].to_numpy(dtype=float) / 1000)))  # converted to KN

            P.append(np.nanmax(abs(df2["CumMassX"].to_numpy(dtype=float) * 9.81 / 1000)))  # converted to KN

            driftX.append(np.nanmax(abs(df3["UX"].to_numpy(dtype=float) * 1000)))
            driftY.append(np.nanmax(abs(df3["UY"].to_numpy(dtype=float) * 1000)))

    # V_x = np.zeros(len(Vx))
    # V_x[0] = Vx[0]
    # for i in range(len(Vx) - 1):
    #     V_x[i + 1] = Vx[i + 1] - Vx[i]
    #
    # V_y = np.zeros(len(Vy))
    # V_y[0] = Vy[0]
    # for i in range(len(Vy) - 1):
    #     V_y[i + 1] = Vy[i + 1] - Vy[i]

    sISFX = 0
    sISFY = 0
    ISFX = 0
    ISFY = 0
    thetaX = 0
    thetaY = 0
    if np.any(Vx):
        sISFX = np.divide(np.multiply(Vx, sh), np.multiply(P, driftX))
        thetaX_ = np.divide(np.multiply(P, driftX), np.multiply(Vx, sh))
        thetaX = np.nanmin(thetaX_)
        ISFX = np.nanmin(sISFX)
    if np.any(Vy):
        sISFY = np.divide(np.multiply(Vy, sh), np.multiply(P, driftY))
        thetaY_ = np.divide(np.multiply(P, driftY), np.multiply(Vy, sh))
        thetaY = np.nanmin(thetaY_)
        ISFY = np.nanmin(sISFY)

    newISFX = np.zeros(len(storyNamelist))
    newISFY = np.zeros(len(storyNamelist))
    newISFX[storyIndex] = sISFX
    newISFY[storyIndex] = sISFY

    resDict = {'resStory': np.flip(np.transpose(np.array([newISFX, newISFY])), axis=0),
               'colName': ["InstabilitySafetyFactorX", "InstabilitySafetyFactorY"],
               'resGlobal': [ISFX, ISFY,thetaX,thetaY],
               'colGlobal': ['ISFX', 'ISFY','thetaX','thetaY']}

    return resDict


# ======================================================================================================================


def get_axialCapacityIndex(SapModel):
    # return axial strength index
    # Axial Demand is the maximum demand due to any load case or combination for columns and shear walls
    # Axial Capacity is calculated base in ACI 318-19 cl 22.4.2

    # allP_df = get_dataFromTable(SapModel, tableKey="Story Forces",
    #                             fieldKeyList=['Story', 'Location', 'P'],
    #                             required_column=['Story', 'P'],
    #                             filterKeyList=['Location'],
    #                             required_row=[['Bottom']])

    allP_df = get_dataFromTable(SapModel, tableKey="Centers Of Mass And Rigidity",
                                fieldKeyList=['Story', 'Diaphragm', 'CumMassX', 'XCM', 'YCM'],
                                required_column=['Story', 'Diaphragm', 'CumMassX', 'XCM', 'YCM'],
                                filterKeyList=None,
                                required_row=None)

    [_, _, storyNames, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    sASI_ = []
    storyIndex = []
    fck = []
    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        # print(story)
        storyP_df = allP_df[allP_df['Story'] == story]
        # storyP_df = allP_df.loc[allP_df['Story'] == story]
        if not storyP_df.size == 0:
            storyP = storyP_df['CumMassX'].to_numpy(dtype=float)
            if np.sum(storyP) == 0:
                sASI_.append(1.0)
                continue

            storyIndex.append(i)
            Pu = np.nanmax(abs(storyP)) * 9.81  # to Newton
            columns = getStoryColumnInfo(SapModel, story=story)
            walls = getStoryWallInfo(SapModel, story=story)

            if columns is None and walls is None:
                sASI_.append(1.0)
                continue

            if walls is None:
                iter_list = columns
            elif columns is None:
                iter_list = walls
            else:
                iter_list = columns + walls

            # colNameList = [cl.name for cl in columns]
            # wallNameList = [wl.name for wl in walls]

            colConcProp = [cl.sectionProp[0] for cl in iter_list]

            Fc_df = get_dataFromTable(SapModel, tableKey="Material Properties - Concrete Data",
                                      fieldKeyList=['Material', 'Fc'],
                                      required_column=['Material', 'Fc'],
                                      filterKeyList=['Material'],
                                      required_row=[colConcProp])

            if not columns is None:
                colRebarProp = [cl.sectionProp[3] for cl in columns]
                Fy_df = get_dataFromTable(SapModel, tableKey="Material Properties - Rebar Data",
                                          fieldKeyList=['Material', 'Fy'],
                                          required_column=['Material', 'Fy'],
                                          filterKeyList=['Material'],
                                          required_row=[colRebarProp])

            cPmax = 0
            wPmax = 0

            if walls is not None:
                il = [w.ilocation for w in walls]
                wName = [w.name for w in walls]

                idf_ = pd.DataFrame(il, columns=['x', 'y', 'z'])
                idf_['name'] = wName
                idf1_ = idf_.drop_duplicates(subset=['x', 'y'])
                unqName = idf1_['name'].values.tolist()

                for i, col in enumerate(walls):
                    if col.name in unqName:
                        Ag = col.area
                        try:
                            Fc = Fc_df.loc[Fc_df['Material'] == col.sectionProp[0]]
                            fc = float(list(set(Fc['Fc']))[0])
                            fck.append(fc)
                        except:
                            pass
                        try:
                            Fy = Fy_df.loc[Fy_df['Material'] == col.sectionProp[3]]
                            fy = float(list(set(Fy['Fy']))[0])
                        except:
                            fy = 413685473.37  # N/m^2

                        wPmax = wPmax + 0.8 * (
                                    0.85 * fc * 0.9975 * Ag + fy * 0.0025 * Ag)  # TODO... 0.0025 minimum considered
            else:
                wPmax = 0

            if columns is not None:
                for i, col in enumerate(columns):
                    Ag = col.area
                    try:
                        Fc = Fc_df.loc[Fc_df['Material'] == col.sectionProp[0]]
                        fc = float(list(set(Fc['Fc']))[0])
                        fck.append(fc)
                    except:
                        pass
                    try:
                        Fy = Fy_df.loc[Fy_df['Material'] == col.sectionProp[3]]
                        fy = float(list(set(Fy['Fy']))[0])
                    except:
                        fy = 413685473.37  # N/m^2

                    cPmax = cPmax + 0.8 * (
                            0.85 * fc * 0.992 * Ag + fy * 0.008 * Ag)  # TODO... 0.8% minimum reinf for column
            else:
                cPmax = 0


            # if walls is not None:
            #     il = [w.ilocation for w in walls]
            #     wName = [w.name for w in walls]
            #
            #     idf_ = pd.DataFrame(il, columns=['x', 'y', 'z'])
            #     idf_['name'] = wName
            #     idf1_ = idf_.drop_duplicates(subset=['x', 'y'])
            #     unqName = idf1_['name'].values.tolist()
            # else:
            #     unqName = []
            #
            # if columns is not None:
            #     cName = [c.name for c in columns]
            #     unqName = unqName + cName
            #
            # for i, col in enumerate(iter_list):
            #     if col.name in unqName:
            #         Ag = col.area
            #         try:
            #             Fc = Fc_df.loc[Fc_df['Material'] == col.sectionProp[0]]
            #             fc = float(list(set(Fc['Fc']))[0])
            #             fck.append(fc)
            #         except:
            #             pass
            #         try:
            #             Fy = Fy_df.loc[Fy_df['Material'] == col.sectionProp[3]]
            #             fy = float(list(set(Fy['Fy']))[0])
            #         except:
            #             fy = 413685473.37  # N/m^2
            #
            #         Pmax = Pmax + 0.8 * (0.85 * fc * 0.98 * Ag + fy * 0.02 * Ag)  #
            #
            #     # print(Pu)

            sASI_.append((cPmax+wPmax) / Pu)
        else:
            sASI_.append(1.0)

    indices = np.where(np.in1d(storyNamelist, storyTower))[0]

    # sASI = np.zeros(len(storyNamelist))
    # sASI[storyIndex] = sASI_

    sASI = sASI_
    ACI = np.nanmean(np.array(sASI_)[indices])

    maxfck = np.nanmax(fck) * 10 ** (-6)
    maxE = 4700 * np.sqrt(maxfck)

    resDict = {'resStory': np.flip(np.array(sASI).reshape(-1, 1), axis=0),
               'colName': ['AxialCapacityIndex'],
               'resGlobal': [maxfck, maxE, ACI],
               'colGlobal': ['Fck', 'E2', 'ACI']}

    return resDict


# ======================================================================================================================
def get_structuralPlanDensityIndex(SapModel):
    # This function computes the Structural Plan Density Index (SPDI) in percentage for each
    # story and return the mean of SPDI.
    [_, _, storyNames, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    # floorArea_df = get_dataFromTable(SapModel, tableKey="Material List by Story",
    #                                  fieldKeyList=['Story', 'FloorArea'],
    #                                  required_column=['Story', 'FloorArea'],
    #                                  filterKeyList=['Story'],
    #                                  required_row=[storyNamelist])
    # floorArea_df = floorArea_df.drop_duplicates()

    # strcolTX = np.zeros((len(storyNamelist)))
    # strcolTY = np.zeros((len(storyNamelist)))
    strcolTX = []
    strcolTY = []
    sSPDI = []
    wallThk = []
    colAreaRatio = []
    wallAreaRatio = []
    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        columns = getStoryColumnInfo(SapModel, story=story)
        walls = getStoryWallInfo(SapModel, story=story)
        # print(story)
        if columns is None and walls is None:
            continue
        # try:
        # storyArea = Storry(SapModel, story).Area
        storyArea = [s.Area for s in stories if s.name == story][0]
        # except:
        #     storyArea = (floorArea_df[floorArea_df['Story'] == story]['FloorArea']).to_numpy(dtype=float).item()

        if walls is None:
            iter_list = columns
        elif columns is None:
            iter_list = walls
        else:
            iter_list = columns + walls

        temp = []
        strArea = 0
        if walls is not None:
            il = [w.ilocation for w in walls]
            wName = [w.name for w in walls]

            idf_ = pd.DataFrame(il, columns=['x', 'y', 'z'])
            idf_['name'] = wName
            idf1_ = idf_.drop_duplicates(subset=['x', 'y'])

            for col in walls:
                if col.name in idf1_['name'].values.tolist():
                    strArea = strArea + col.area
                    temp.append(col.thickness)
            wallThk.append(np.nanmax(temp) * 1000)
            wallAreaRatio.append(strArea / storyArea * 100)  # in percentage
        else:
            wallThk.append(0)
            wallAreaRatio.append(0)

        colA = []
        tX = []
        tY = []
        strArea = 0
        # print(story,storyArea)

        if columns is not None:

            for col in columns:
                tX.append(col.tX)
                tY.append(col.tY)
                colA.append(col.area)
                strArea = strArea + col.area
            ind = np.argmax(colA)

            if type(ind) == np.ndarray: ind = ind[0]
            strcolTX.append(tX[ind] * 1000)
            strcolTY.append(tY[ind] * 1000)
            colAreaRatio.append(strArea / storyArea * 100)  # in percentage
        else:
            strcolTX.append(0)
            strcolTY.append(0)
            colAreaRatio.append(0)  # in percentage

        # strArea = 0
        # for col in iter_list:
        #     strArea = strArea + col.area
    sSPDI = np.array(wallAreaRatio) + np.array(colAreaRatio)  # in percentage
    sSPDI = sSPDI.tolist()
    indices = np.where(np.in1d(storyNamelist, storyTower))[0]

    # SPDI = np.round(np.nanmean(np.array(sSPDI)[np.array(sSPDI)<40]), 3)

    col_ = np.array(colAreaRatio)[indices]
    wal_ = np.array(wallAreaRatio)[indices]

    CPDI = np.round(np.nanmean(col_[col_ < 40]), 3)
    WPDI = np.round(np.nanmean(wal_[wal_ < 40]), 3)
    SPDI = CPDI + WPDI

    wallThickness = round(np.nanmax(wallThk), 2)
    colX = round(np.nanmax(strcolTX), 2)
    colY = round(np.nanmax(strcolTY), 2)

    resDict = {'resStory': np.flip(np.transpose(np.array([strcolTX, strcolTX, wallThk,
                                                          colAreaRatio, wallAreaRatio, sSPDI])), axis=0),
               'colName': ['colThkX_mm', 'colThkX_mm', 'avgWallThickness', "colAreaRatio_%",
                           "wallAreaRatio_%", "StructuralPlanDensityIndex_%"],
               'resGlobal': [colX, colY, wallThickness, CPDI.item(), WPDI.item(), SPDI.item()],
               'colGlobal': ['colX_mm', 'colY_mm', 'wallThickness_mm', 'CPDI_%', 'WPDI_%', 'SPDI_%']}
    return resDict


# ======================================================================================================================

def get_shorteningIndex(SapModel):
    # return Shortening Index (SI)
    # Shortening Index (SI) is the maximum vertical displacement of any floor divided by story height
    # shortening index is calculated for all story but this function only the maximum among all the story
    # which can be modified later if necessary

    ret = SapModel.SetPresentUnits(10)
    [_, storyNumber, storyNames, _, H, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    storyMaxDispZ = []

    # [_, ptrnName, ret] = SapModel.LoadPatterns.GetNameList()
    [_, caseName, ret] = SapModel.LoadCases.GetNameList()
    try:
        [_, comboName, ret] = SapModel.RespCombo.GetNameList()
    except:
        comboName = None
    selCase = []
    otherCases = []
    for case in caseName:
        [caseType, subType, dType, _, _, ret] = SapModel.LoadCases.GetTypeOAPI_1(case)
        if dType in [1, 2, 3, 4, 11] and caseType != 10:
            selCase.append(case)
        if not caseType in [3, 10]:
            otherCases.append(case)

    ret = SapModel.DatabaseTables.SetLoadCasesSelectedForDisplay(selCase)
    ret = SapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay([])

    ALLjointDispZ_df = get_dataFromTable(SapModel, tableKey="Joint Displacements",
                                         fieldKeyList=['Story', 'Uz'],
                                         required_column=['Story', 'Uz'],
                                         filterKeyList=None,
                                         required_row=None)

    ret = SapModel.DatabaseTables.SetLoadCasesSelectedForDisplay(otherCases)
    try:
        ret = SapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay(comboName)
    except:
        pass

    # ALLjointDispZ_df = pd.read_csv("tempdisp.csv")

    storyHeights = []
    for i, storyName in enumerate(storyNamelist):
        jointDispZ_df = ALLjointDispZ_df.loc[ALLjointDispZ_df['Story'] == storyName]
        jointDispZ = jointDispZ_df['Uz'].to_numpy(dtype=float)  # converted to meter
        if not jointDispZ.size == 0:
            storyMaxDispZ.append(np.nanmax(abs(jointDispZ)))
            storyHeights.append(H[i])

    ALLjointDispZ_df = None

    storyMaxDispZ = np.array(storyMaxDispZ).reshape(-1, 1)
    storyHeights = np.array(storyHeights).reshape(-1, 1)
    sSI = np.divide(storyMaxDispZ, storyHeights) * 100  # in percentage

    indices = np.where(np.in1d(storyNamelist, storyTower))[0]
    SI = np.nanmax(np.array(sSI)[indices])

    # plt.plot(SI,np.linspace(0,storyNumber,storyNumber))
    # plt.show()

    # return storyMaxDispZ, sSI, SI
    storyMaxDispZ = storyMaxDispZ * 1000  # converted to mm
    resDict = {'resStory': np.flip(np.concatenate([storyMaxDispZ, sSI], axis=1), axis=0),
               'colName': ["StoryMaxDispZ_mm", "ShorteningIndex_%"],
               'resGlobal': [SI.item()],
               'colGlobal': ['SI_%']}

    return resDict


# ======================================================================================================================


def timePeriodManual(SapModel):
    # diaphragmMass_df = get_dataFromTable(SapModel, tableKey="Mass Summary by Diaphragm",
    #                                      fieldKeyList=['Story', 'MassX'],
    #                                      required_column=['MassX'],
    #                                      filterKeyList=None,
    #                                      required_row=None)
    #
    # m = diaphragmMass_df['MassX'].to_numpy(dtype=float)

    diaphragmMass_df = get_dataFromTable(SapModel, tableKey="Mass Summary by Story",
                                         fieldKeyList=['Story', 'UX'],
                                         required_column=['Story', 'UX'],
                                         filterKeyList=None,
                                         required_row=None)

    m = diaphragmMass_df['UX'][0:-1].to_numpy(dtype=float)

    [_, _, storyNames, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    StiffX = []
    StiffY = []

    storyStiffness_df = get_dataFromTable(SapModel, tableKey="Story Stiffness",
                                          fieldKeyList=['Story', 'OutputCase', 'StiffX', 'StiffY'],
                                          required_column=['Story', 'OutputCase', 'StiffX', 'StiffY'],
                                          filterKeyList=None,
                                          required_row=None)

    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        s_df = storyStiffness_df.loc[storyStiffness_df['Story'] == story]
        if not s_df.size == 0:
            stiffX = np.nanmax(s_df['StiffX'].to_numpy(dtype=float))
            stiffY = np.nanmax(s_df['StiffY'].to_numpy(dtype=float))
            StiffX.append(stiffX)
            StiffY.append(stiffY)
    kx = np.flip(StiffX)
    ky = np.flip(StiffY)

    # m = [547460, 657780, 657780, 657780, 657780]
    # k = [600, 1200, 1800]

    nDOF = len(m)

    M = np.zeros((nDOF, nDOF))
    for i in range(nDOF):
        M[i, i] = m[i]
    # print(M)
    KX = np.zeros((nDOF, nDOF))
    KY = np.zeros((nDOF, nDOF))

    for i in range(nDOF):
        if not (i == 0 or i == (nDOF - 1)):
            KX[i, i] = kx[i - 1] + kx[i]
            KX[i, i - 1] = -kx[i - 1]
            KX[i, i + 1] = -kx[i]

            KY[i, i] = ky[i - 1] + ky[i]
            KY[i, i - 1] = -ky[i - 1]
            KY[i, i + 1] = -ky[i]


        else:
            if i == 0:
                KX[i, i] = kx[i]
                KX[i, i + 1] = -kx[i]

                KY[i, i] = ky[i]
                KY[i, i + 1] = -ky[i]

            elif i == (nDOF - 1):
                KX[i, i] = kx[i - 1] + kx[i]
                KX[i, i - 1] = -kx[i - 1]

                KY[i, i] = ky[i - 1] + ky[i]
                KY[i, i - 1] = -ky[i - 1]

    wx, vx = np.linalg.eig(np.dot(KX, np.linalg.inv(M)))
    wy, vy = np.linalg.eig(np.dot(KY, np.linalg.inv(M)))
    Tx = 2 * np.pi / np.sqrt(wx)
    Ty = 2 * np.pi / np.sqrt(wy)
    return Tx, Ty


# ======================================================================================================================


def get_torsionalIrregularityIndex(SapModel):
    # Torsional Irregularity Index i.e ratio of maximum over average story displacement is directly
    # obtained from ETABS data table.
    [_, _, storyNames, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    sTIIX_ = []
    sTIIY_ = []
    ratio_df = get_dataFromTable(SapModel, tableKey="Diaphragm Max Over Avg Drifts",
                                 fieldKeyList=['Story', 'Item', 'Ratio'],
                                 required_column=['Story', 'Item', 'Ratio'],
                                 filterKeyList=None,
                                 required_row=None)

    storyIndex = []
    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        s_df = ratio_df.loc[ratio_df['Story'] == story]
        if not s_df.size == 0:
            storyIndex.append(i)
            s_df = s_df.set_index("Item")
            x_df = s_df.filter(like="X", axis=0)
            y_df = s_df.filter(like="Y", axis=0)

            if not x_df.empty:
                stiix = np.nanmax(x_df['Ratio'].to_numpy(dtype=float))
                sTIIX_.append(stiix)
            else:
                sTIIX_.append(0.0)
            if not y_df.empty:
                stiiy = np.nanmax(y_df['Ratio'].to_numpy(dtype=float))
                sTIIY_.append(stiiy)
            else:
                sTIIY_.append(0.0)
        else:
            sTIIX_.append(0.0)
            sTIIY_.append(0.0)

    indices = np.where(np.in1d(storyNamelist, storyTower))[0]

    TIIX = np.nanmean(np.array(sTIIX_)[indices])
    TIIY = np.nanmean(np.array(sTIIY_)[indices])

    # sTIIX = np.zeros(len(storyNamelist))
    # sTIIY = np.zeros(len(storyNamelist))
    # sTIIX[storyIndex] = sTIIX_
    # sTIIY[storyIndex] = sTIIY_

    # return TIIX, TIIY
    resDict = {'resStory': np.flip(np.transpose(np.array([sTIIX_, sTIIY_])), axis=0),
               'colName': ["TorsionalIrregularityX", "TorsionalIrregularityIndexY"],
               'resGlobal': [TIIX.item(), TIIY.item()],
               'colGlobal': ['TIIX', 'TIIY']}

    return resDict


# ======================================================================================================================

def get_InherentTorsionRatio(SapModel):
    [_, _, storyNames, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    sITRX_ = []
    sITRY_ = []
    centroid_df = get_dataFromTable(SapModel, tableKey="Centers Of Mass And Rigidity",
                                    fieldKeyList=['Story', 'Diaphragm', 'MassX', 'XCM', 'YCM', 'XCR', 'YCR'],
                                    required_column=['Story', 'Diaphragm', 'MassX', 'XCM', 'YCM', 'XCR', 'YCR'],
                                    filterKeyList=None,
                                    required_row=None)

    storyIndex = []
    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        c_df = centroid_df.loc[centroid_df['Story'] == story]
        c_df = c_df.dropna()
        if not c_df.size == 0 and sum(c_df['MassX'].to_numpy(dtype=float)) != 0:
            storyIndex.append(i)

            XCM = np.dot(c_df['MassX'].to_numpy(dtype=float), c_df['XCM'].to_numpy(dtype=float)) / sum(
                c_df['MassX'].to_numpy(dtype=float))

            YCM = np.dot(c_df['MassX'].to_numpy(dtype=float), c_df['YCM'].to_numpy(dtype=float)) / sum(
                c_df['MassX'].to_numpy(dtype=float))

            XCR = np.dot(c_df['MassX'].to_numpy(dtype=float), c_df['XCR'].to_numpy(dtype=float)) / sum(
                c_df['MassX'].to_numpy(dtype=float))

            YCR = np.dot(c_df['MassX'].to_numpy(dtype=float), c_df['YCR'].to_numpy(dtype=float)) / sum(
                c_df['MassX'].to_numpy(dtype=float))

            ex = np.abs(XCM - XCR)
            ey = np.abs(YCM - YCR)

            # s = Storry(SapModel, story)
            LX = [s.LenX for s in stories if s.name == story][0]
            LY = [s.LenY for s in stories if s.name == story][0]

            # LX = s.LenX
            # LY = s.LenY
            # s = None
            sITRX_.append(ex / LX)
            sITRY_.append(ey / LY)
        else:
            sITRX_.append(0.0)
            sITRY_.append(0.0)

    indices = np.where(np.in1d(storyNamelist, storyTower))[0]
    ITRX = np.nanmean(np.array(sITRX_)[indices])
    ITRY = np.nanmean(np.array(sITRY_)[indices])

    # sITRX = np.zeros(len(storyNamelist))
    # sITRY = np.zeros(len(storyNamelist))
    # sITRX[storyIndex] = sITRX_
    # sITRY[storyIndex] = sITRY_

    resDict = {'resStory': np.flip(np.transpose(np.array([sITRX_, sITRY_])), axis=0),
               'colName': ["InherentTorsionRatioX", "InherentTorsionRatioY"],
               'resGlobal': [ITRX.item(), ITRY.item()],
               'colGlobal': ['ITRX', 'ITRY']}

    return resDict


# ======================================================================================================================

def get_storyStiffnessIrregularity(SapModel):
    [_, _, storyNames, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    StiffX = []
    StiffY = []

    try:
        storyStiffness_df = get_dataFromTable(SapModel, tableKey="Story Stiffness",
                                              fieldKeyList=['Story', 'OutputCase', 'StiffX', 'StiffY'],
                                              required_column=['Story', 'OutputCase', 'StiffX', 'StiffY'],
                                              filterKeyList=None,
                                              required_row=None)
    except:
        storyStiffness_df = get_dataFromTable(SapModel, tableKey="Story Stiffness",
                                              fieldKeyList=['Story', 'OutputCase', 'StiffXh', 'StiffYh'],
                                              required_column=['Story', 'OutputCase', 'StiffXh', 'StiffYh'],
                                              filterKeyList=None,
                                              required_row=None)
        storyStiffness_df[['StiffX', 'StiffY']] = storyStiffness_df[['StiffXh', 'StiffYh']]
    temp = []
    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        s_df = storyStiffness_df.loc[storyStiffness_df['Story'] == story]
        if not s_df.size == 0:
            # s_df = s_df.set_index("OutputCase")
            # x_df = s_df.filter(like="X", axis=0)
            # y_df = s_df.filter(like="Y", axis=0)
            temp.append(story)
            stiffX = np.nanmax(s_df['StiffX'].to_numpy(dtype=float))
            stiffY = np.nanmax(s_df['StiffY'].to_numpy(dtype=float))
            StiffX.append(stiffX)
            StiffY.append(stiffY)

    storyNamelist = temp
    # StiffX = [k for k in StiffX if k != 0.0]
    # StiffY = [k for k in StiffY if k != 0.0]

    SSIx1 = []
    SSIx2 = []
    SSIy1 = []
    SSIy2 = []
    for i in range(len(StiffX)):
        try:
            SSIx1.append(StiffX[i] / StiffX[i + 1])
            SSIy1.append(StiffY[i] / StiffY[i + 1])
        except:
            SSIx1.append(1.0)
            SSIy1.append(1.0)
        try:
            SSIx2.append(3 * StiffX[i] / (StiffX[i + 1] + StiffX[i + 2] + StiffX[i + 3]))
            SSIy2.append(3 * StiffY[i] / (StiffY[i + 1] + StiffY[i + 2] + StiffY[i + 3]))
        except:
            SSIx2.append(1.0)
            SSIy2.append(1.0)

    # sSSIX = np.amax([np.array(SSIx1), np.array(SSIx2)])
    # sSSIY = np.amax([np.array(SSIy1), np.array(SSIy2)])

    SSIx1 = np.flip(np.array(SSIx1), axis=0).reshape(-1, 1)
    SSIy1 = np.flip(np.array(SSIy1), axis=0).reshape(-1, 1)

    SSIx1 = np.nan_to_num(SSIx1, nan=1, posinf=1, neginf=1)
    SSIy1 = np.nan_to_num(SSIy1, nan=1, posinf=1, neginf=1)

    indices = np.where(np.in1d(storyNamelist, storyTower))[0]

    SSIX = np.nanmax([np.nanmax(np.array(SSIx1)[indices]), np.nanmax(np.array(SSIx2)[indices])])
    SSIY = np.nanmax([np.nanmax(np.array(SSIy1)[indices]), np.nanmax(np.array(SSIy2)[indices])])

    # return SSIX, SSIY

    resDict = {'resStory': np.concatenate([SSIx1, SSIy1], axis=1),
               'colName': ["StoryStiffnessIrregularityX", "StoryStiffnessIrregularityY"],
               'resGlobal': [SSIX.item(), SSIY.item()],
               'colGlobal': ['SSIX', 'SSIY']}

    return resDict


# ======================================================================================================================

def get_storyMassIrregularity(SapModel):
    [_, _, storyNames, elev, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    sMass = []

    storyMass_df = get_dataFromTable(SapModel, tableKey="Mass Summary by Story",
                                     fieldKeyList=['Story', 'UX'],
                                     required_column=['Story', 'UX'],
                                     filterKeyList=None,
                                     required_row=None)

    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        s_df = storyMass_df.loc[storyMass_df['Story'] == story]
        if s_df.size != 0:
            sMass.append(s_df['UX'].to_numpy(dtype=float))
        else:
            sMass.append(0.0)

    massIR = []
    # sMass = [k for k in sMass if k != 0.0]
    for i in range(len(sMass)):
        m1 = np.array([1.0])
        m2 = np.array([1.0])
        try:
            if i != (len(sMass) - 1):
                m1 = sMass[i] / sMass[i + 1]
            if i != 0:
                m2 = sMass[i] / sMass[i - 1]
            massIRR = np.nanmax([m1, m2])
            # if massIRR >= 1:
            #     massIRR = 1/massIRR
            # else:
            #     massIRR = massIRR

            massIR.append(massIRR)
        except:
            massIR.append(1.0)

    indices = np.where(np.in1d(storyNamelist, storyTower))[0]

    massIR = np.nan_to_num(massIR, nan=1, posinf=1, neginf=1)

    MII = np.nanmean(np.array(massIR)[indices])

    # return MII
    massIR = np.flip(massIR, axis=0).reshape(-1, 1)
    resDict = {'resStory': massIR,
               'colName': ["StoryMassIrregularityX"],
               'resGlobal': [MII.item()],
               'colGlobal': ['MII']}

    return resDict


# ======================================================================================================================

def get_vibrationSeriviceabilityIndex(SapModel):
    [_, _, storyNames, _, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()
    # accX = []
    # accY = []
    # accZ = []
    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    acc = []
    try:
        storyAcc_df = get_dataFromTable(SapModel, tableKey="Story Accelerations",
                                        fieldKeyList=['Story', 'OutputCase', 'UX', 'UY', 'UZ'],
                                        required_column=['Story', 'OutputCase', 'UX', 'UY', 'UZ'],
                                        filterKeyList=None,
                                        required_row=None)
    except:
        return None
    indices = np.where(np.in1d(storyNamelist, storyTower))[0]

    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        if i in indices:
            s_df = storyAcc_df.loc[storyAcc_df['Story'] == story]
            if not s_df.size == 0:
                ux = np.nanmax(s_df['UX'].to_numpy(dtype=float))
                uy = np.nanmax(s_df['UY'].to_numpy(dtype=float))
                uz = np.nanmax(s_df['UZ'].to_numpy(dtype=float))
                # accX.append(ux/9806.65)
                # accY.append(uy/9806.65)
                # accZ.append(uz/9806.65)
                acc.append(np.nanmax([(ux / 9.80665)*1000, (uy / 9.80665)*1000, (uz / 9.80665)*1000]))  # UNIT mili-g
            else:
                acc.append(0.0)

    acc_ = np.zeros((len(storyNamelist), 1))
    acc_[indices, 0] = acc
    VSI = np.nanmax(acc)
    # return VSI
    resDict = {'resStory': np.flip(np.array(acc_).reshape(-1, 1), axis=0),
               'colName': ['VibrationServiceabilityIndex'],
               'resGlobal': [VSI.item()],
               'colGlobal': ['VSI']}

    return resDict


# ======================================================================================================================

def get_perimeterIndexes(SapModel):
    # This function returns two important parameter of building related to perimeter.
    # 1. Perimeter Enclosure Factor (PEF) = baseline perimeter/ actual perimeter
    # 2. Perimeter Area Index (PAI) = area around the perimeter/ floor area
    [_, _, storyNames, elev, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    for story in storyNames:
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
        except:
            pass

    sPEF = []
    sPAI = []
    indices = np.where(np.in1d(storyNamelist, storyTower))[0]
    # indices = [8, 9, 10]
    for i, story in enumerate(storyNamelist):  # i in range(1):   #
        if i in indices:
            try:
                # s = Storry(SapModel, story)
                # area = s.Area
                # Lp = s.perimeter
                area = [s.Area for s in stories if s.name == story][0]
                Lp = [s.perimeter for s in stories if s.name == story][0]

                if area == 0.0:
                    sPEF.append(0.0)
                    sPAI.append(0.0)
                    continue
                if Lp == 0.0:
                    sPEF.append(0.0)
                    sPAI.append(0.0)
                    continue
                Bp = 2 * np.sqrt(np.pi * area)
                sPEF.append(Bp / Lp)
                sPAI.append(Lp * 4.5 / area)
            except:
                sPEF.append(0.0)
                sPAI.append(0.0)
                continue

    sPEF_ = np.zeros((len(storyNamelist)))
    sPEF_[indices] = sPEF

    sPAI_ = np.zeros((len(storyNamelist)))
    sPAI_[indices] = sPAI

    PEF = np.nanmedian(sPEF)
    PAI = np.nanmedian(sPAI)

    # return PEF, PAI
    resDict = {'resStory': np.flip(np.transpose(np.array([sPEF_, sPAI_])), axis=0),
               'colName': ["PerimeterEnclosureFactor", "PerimeterAreaIndex"],
               'resGlobal': [PEF.item(), PAI.item()],
               'colGlobal': ['PEF', 'PAI']}

    return resDict


# ======================================================================================================================


def gust_effect_factor(SapModel, h=183, B=30, L=30, n1=0.2, V=51, beta=0.01, exposure="B"):
    from math import sqrt, log, exp
    # This program calculates gust response factor as per ASCE 7-16
    # ========= inputs ==============================================
    [baseElev, _, storyNames, storyElev_, _, _, _, _, _, _, ret] = SapModel.Story.GetStories_2()

    # empty stories are removed from the list
    storyNamelist = []
    storyElev = []
    for i, story in enumerate(storyNames):
        try:
            [_, _, ret] = SapModel.PointObj.GetNameListOnStory(story)
            storyNamelist.append(story)
            storyElev.append(storyElev_[i])
        except:
            pass

    # s1 = Storry(SapModel, storyNamelist[1])
    # B = np.nanmin([s1.LenX, s1.LenY])
    # L = np.nanmax([s1.LenX, s1.LenY])

    indices = np.where(np.in1d(storyNamelist, storyTower))[0]
    story = storyNamelist[indices[0]]
    lx = [s.LenX for s in stories if s.name == story][0]
    ly = [s.LenY for s in stories if s.name == story][0]

    B = np.nanmin([lx, ly])
    L = np.nanmax([lx, ly])

    df = get_dataFromTable(SapModel, tableKey="Modal Participating Mass Ratios",
                           fieldKeyList=['Mode', "Period"],
                           required_column=['Mode', "Period"],
                           filterKeyList=None,
                           required_row=None)

    n1 = np.nanmax(df['Period'].to_numpy(dtype='float'))

    df = get_dataFromTable(SapModel, tableKey="Load Pattern Definitions",
                           fieldKeyList=['Type', "AutoLoad"],
                           required_column=['Type', "AutoLoad"],
                           filterKeyList=None,
                           required_row=None)

    if df[df['Type'] == "Wind"].empty:
        return None

    autoLoad = df[(df['Type'] == "Wind") & (df['AutoLoad'] != "User Loads")]['AutoLoad']
    if autoLoad.empty:
        return None
    autoLoad = autoLoad.iloc[0]
    windTable = "Load Pattern Definitions - Auto Wind - " + autoLoad

    df = get_dataFromTable(SapModel, tableKey=windTable,
                           fieldKeyList=['WindSpeed', "ExpType", "GustFact", "TopStory", "BotStory"],
                           required_column=['WindSpeed', "ExpType", "GustFact", "TopStory", "BotStory"],
                           filterKeyList=None,
                           required_row=None)
    V = df['WindSpeed'].to_numpy(dtype="float")[0] * 0.44704  # converted to m/s
    exposure = df['ExpType'][0]
    gustFact = df['GustFact'].to_numpy(dtype="float")[0]

    topStory = df['TopStory'][0]
    botStory = df['BotStory'][0]

    storyElev = np.array(storyElev)
    storyNamelist = np.array(storyNamelist)
    if botStory in storyNamelist:
        botElev = storyElev[storyNamelist == botStory]
    else:
        botElev = baseElev

    if topStory in storyNamelist:
        topElev = storyElev[storyNamelist == topStory]
    else:
        topElev = np.nanmax(storyElev)
    h = topElev - botElev

    # =========================================================================================================
    # Table 26.11-1 Terrain Exposure Constants (SI Units)
    terrain_exposure_constants = {"B": [7.0, 365.76, 1 / 7, 0.84, 1 / 4.0, 0.45, 0.3, 97.54, 1 / 3.0, 9.14],
                                  "C": [9.5, 274.32, 1 / 9.5, 1.00, 1 / 6.5, 0.65, 0.20, 152.40, 1 / 5.0, 4.57],
                                  "D": [11.5, 213.36, 1 / 11.5, 1.07, 1 / 9.0, 0.80, 0.15, 198.12, 1 / 8.0, 2.13]}

    # inputs = input("Enter inputs as vector [h, b, L , n1, V, beta]: ")
    #
    #
    # h, B, L , n1, V, beta = [float(val) for val in inputs.split(",")]
    # h = 10          # height of structure
    # B = 20
    # L = 30
    # n1  = 0.6       # fundamental natural frequency
    # V = 20          # basic wind speed (m/s)
    # beta = 0.05     # damping ratio

    z_bar = 0.6 * h  # equivalent height of structure (=0.6*h)
    alpha_bar = terrain_exposure_constants[exposure][4]
    b_bar = terrain_exposure_constants[exposure][5]
    c = terrain_exposure_constants[exposure][6]
    l = terrain_exposure_constants[exposure][7]
    e_bar = terrain_exposure_constants[exposure][8]
    Iz = c * (10 / z_bar) ** (1 / 6)  # Intensity_of_turbulence
    gQ = 3.4  # peak_factor_for_backgroundresponse
    gv = 3.4  # peak_factor_for_windresponse
    gR = sqrt(2 * log(3600 * n1)) + 0.577 / (sqrt(2 * log(3600 * n1)))  # peak_factor_for_resonantresponse
    Lz = l * (z_bar / 10) ** e_bar
    Q = sqrt(1 / (1 + 0.63 * ((B + h) / Lz) ** 0.63))  # background_response_factor
    Vz = b_bar * (z_bar / 10) ** alpha_bar * V
    N1 = n1 * Lz / Vz
    Rn = 7.47 * N1 / (1 + 10.3 * N1) ** (5 / 3)

    nus = [4.6 * n1 * h / Vz, 4.6 * n1 * B / Vz, 15.4 * n1 * L / Vz]
    R_l = [1 / nu - 1 / (2 * nu ** 2) * (1 - exp(-2 * nu)) for nu in nus]
    R = sqrt(1 / beta * Rn * R_l[0] * R_l[1] * (0.53 + 0.47 * R_l[2]))  # resonant_response_factor
    gust_effect_factor_flexible = 0.925 * (1 + 1.7 * Iz * sqrt(gQ * gQ * Q * Q + gR * gR * R * R)) / (1 + 1.7 * gv * Iz)
    cache = (z_bar, alpha_bar, b_bar, c, Iz, Vz, Q * Q, Rn, R_l[0], R_l[1], R_l[2], nus, gR)

    resDict = {'resStory': None,
               'resGlobal': [gust_effect_factor_flexible.item(), gustFact],
               'colGlobal': ['GustEffectFactorFlexible', 'GustEffectFactorCode']}

    return resDict


# ======================================================================================================================


def get_carbonEmission(SapModel):
    MPa = np.array([18, 21, 24, 27, 30, 35, 36, 40, 43, 45, 47, 48, 50, 60, 65])
    CO2 = np.array([230, 246.75, 290, 305.14, 355.6, 352.7, 352.7, 371.42, 369, 370,
                    367.9, 444.3, 385.6, 391.94, 453.84])

    from scipy.interpolate import make_interp_spline
    spl = make_interp_spline(MPa, CO2, k=1)

    try:
        df = get_dataFromTable(SapModel, tableKey="Material List by Object Type",
                               fieldKeyList=['ObjectType', "Material", 'Weight'],
                               required_column=['ObjectType', "Material", 'Weight'],
                               filterKeyList=None,
                               required_row=None)

        matName = df['Material']
        matCO2 = []
        E = []
        for i, mat in enumerate(matName):
            [matType, symType, ret] = SapModel.PropMaterial.GetTypeOAPI(mat)
            if matType == 2:
                [e, u, a, g, ret] = SapModel.PropMaterial.GetMPIsotropic(mat)
                E.append(e * 10 ** (-6))
                fck = ((e * 10 ** (-6)) / 4700) ** 2
                co2 = spl(fck)
                concVol = (float(df.iloc[i, 2]) / 9.81) / 2400.68  # m^3

                matCO2.append(concVol * co2)
        totalMass_kg = sum(df.iloc[:, 2].astype(float)) / 9.81
        maxE = np.nanmax(E)  # in MPa
        totalCO2 = np.nansum(matCO2)

    except:
        df = get_dataFromTable(SapModel, tableKey="Mass Summary by Group",
                               fieldKeyList=['Group', 'SelfWeight'],
                               required_column=['Group', 'SelfWeight'],
                               filterKeyList=None,
                               required_row=None)

        totalMass_kg = np.nanmax(df['SelfWeight'].to_numpy(dtype=float)) / 9.81  # kg

        [_, matName, ret] = SapModel.PropMaterial.GetNameList()
        E = []
        for mat in matName:
            [matType, symType, ret] = SapModel.PropMaterial.GetTypeOAPI(mat)
            if matType == 2:
                [e, u, a, g, ret] = SapModel.PropMaterial.GetMPIsotropic(mat)
                E.append(e * 10 ** (-6))
        maxE = np.nanmax(E)  # in MPa

        fck = (maxE / 4700) ** 2
        co2 = spl(fck)

        concVol = totalMass_kg / 2400.68  # m^3

        totalCO2 = concVol * co2

    emissionFactor = totalCO2 / totalMass_kg

    resDict = {'resStory': None,
               'resGlobal': [totalCO2, emissionFactor],
               'colGlobal': ['totalCO2_kg', 'emissionFactor']}

    return resDict


# =======================================================================================================================

def podiumLevelPlot(figsize=(8, 6), xlabel="", ylabel="Story", title="Story Respose Plot"):
    import matplotlib.pyplot as plt
    from sklearn.cluster import KMeans
    from sklearn.decomposition import PCA

    path = "E:\\00_Thesis Project\\RUN\\storyResults"

    names = []
    roots = []
    files = []

    for root, dira, file in os.walk(path):
        for filename in file:
            if filename.endswith(".xlsx"):
                names.append(filename)
                roots.append(root)
                files.append(os.path.join(root, filename))

    plt.style.use('dark_background')
    plt.rcParams['figure.figsize'] = figsize
    SMALL_SIZE = 12
    MEDIUM_SIZE = 14
    BIGGER_SIZE = 18

    plt.rc('font', family='Arial', size=SMALL_SIZE)  # controls default text sizes
    plt.rc('axes', titleweight='bold', titlesize=BIGGER_SIZE)  # fontsize of the axes title
    plt.rc('axes', labelweight='bold', labelsize=MEDIUM_SIZE)  # fontsize of the x and y labels
    plt.rc('xtick', labelsize=SMALL_SIZE)  # fontsize of the tick labels
    plt.rc('ytick', labelsize=SMALL_SIZE)  # fontsize of the tick labels
    #     plt.rc('legend', fontsize=SMALL_SIZE)  # legend fontsize
    #     plt.rc('figure', titleweight = 'bold',titlesize=BIGGER_SIZE)  # fontsize of the figure title

    fig, ax = plt.subplots()

    i = 0
    maxH = 0
    for file in files:
        try:
            data = pd.read_excel(file, sheet_name="resStory")
            data = data.dropna()
            df = data.iloc[:, 1:6].copy()

            kmeans = KMeans(n_clusters=3, random_state=0)

            df['cluster'] = kmeans.fit_predict(df)

            pca = PCA(2)
            df['x'] = pca.fit_transform(df)[:, 0]
            df['y'] = pca.fit_transform(df)[:, 1]
            df = df.reset_index()

            story_cluster = df[['cluster', 'x', 'y']].copy()
            story_cluster['Story'] = data['Story'].copy()

            # story_cluster = story_cluster.set_index('Story')

            temp = np.flip(story_cluster['cluster'].to_numpy())

            temp3 = np.argmax((np.sum(temp == 0), np.sum(temp == 1), np.sum(temp == 2)))

            temp2 = np.where(temp == temp3)[0][0]
            ind = len(temp) - temp2

            podNum = temp2
            cumH = np.cumsum(df['StoryHeight_m'])
            podH = cumH[podNum]
            totH = np.max(cumH)

            if totH > maxH:
                maxH = totH

            podium_level = story_cluster['Story'].iloc[ind - 1]

            groups = story_cluster.groupby("cluster")
            grp = [g for n, g in groups]
            storyTower = grp[temp3]['Story'].to_list()

            strPoints = pd.DataFrame()
            strPoints['x'] = [i, i]
            strPoints['y'] = [0, totH]

            ax.plot(strPoints['x'], strPoints['y'], alpha=0.7)
            ax.scatter(i, podH, marker='_')
            i = i + 1
        except:
            pass

    ax.spines['left'].set_position('zero')
    ax.spines['right'].set_color('none')
    ax.spines['bottom'].set_position('zero')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_bounds(0, maxH)
    ax.spines['bottom'].set_bounds(0, i)
    ax.xaxis.set_ticks_position('bottom')
    ax.yaxis.set_ticks_position('left')

    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    plt.show()


# =======================================================================================================================

def storywiseplot(plotColumn=None, FUNC='max', lb=None, ub=None, figsize=(6, 8),
                  xlabel="x", ylabel="Story (%)", title="Story Respose Plot"):
    import matplotlib.pyplot as plt

    if plotColumn is None:
        plotColumn = ["DriftX_%", "DriftY_%"]

    path = "E:\\00_Thesis Project\\RUN\\storyResults"

    names = []
    roots = []
    files = []

    for root, dira, file in os.walk(path):
        for filename in file:
            if filename.endswith(".xlsx"):
                names.append(filename)
                roots.append(root)
                files.append(os.path.join(root, filename))

    plt.style.use('dark_background')
    plt.rcParams['figure.figsize'] = figsize
    SMALL_SIZE = 12
    MEDIUM_SIZE = 14
    BIGGER_SIZE = 18

    plt.rc('font', family='Arial', size=SMALL_SIZE)  # controls default text sizes
    plt.rc('axes', titleweight='bold', titlesize=BIGGER_SIZE)  # fontsize of the axes title
    plt.rc('axes', labelweight='bold', labelsize=MEDIUM_SIZE)  # fontsize of the x and y labels
    plt.rc('xtick', labelsize=SMALL_SIZE)  # fontsize of the tick labels
    plt.rc('ytick', labelsize=SMALL_SIZE)  # fontsize of the tick labels
    #     plt.rc('legend', fontsize=SMALL_SIZE)  # legend fontsize
    #     plt.rc('figure', titleweight = 'bold',titlesize=BIGGER_SIZE)  # fontsize of the figure title

    fig, ax = plt.subplots()
    linealpha = 0.35

    xmin = 0
    xmax = 0
    for file in files:
        # print(file)
        try:
            data = pd.read_excel(file, sheet_name="resStory")
            data = data.dropna()
            df = data[plotColumn].copy()
            nStory = df.shape[0]

            if len(plotColumn) > 1:
                if FUNC == 'max':
                    x = df.max(axis=1)
                elif FUNC == 'min':
                    x = df.min(axis=1)
                elif FUNC == 'mean':
                    x = df.mean(axis=1)
            else:
                x = df.squeeze()

            if x.min() < xmin:
                xmin = x.min()
            if x.max() > xmax:
                xmax = x.max()

            y = np.linspace(100, 0, nStory)

            if (lb is not None) and (ub is not None):
                xmin = lb
                xmax = ub
                if (np.nanmax(x) < ub) and (np.nanmax(x) > lb):
                    ax.plot(x, y, alpha=linealpha)

            elif ub is not None:
                xmax = ub
                if np.nanmax(x) < ub:
                    ax.plot(x, y, alpha=linealpha)

            elif lb is not None:
                xmin = lb
                if np.nanmax(x) > lb:
                    ax.plot(x, y, alpha=linealpha)
            else:
                ax.plot(x, y, alpha=linealpha)
        except:
            pass

    ax.spines['left'].set_position('zero')
    ax.spines['right'].set_color('none')
    ax.spines['bottom'].set_position('zero')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_bounds(y.min(), y.max())
    ax.spines['bottom'].set_bounds(xmin, xmax)
    ax.xaxis.set_ticks_position('bottom')
    ax.yaxis.set_ticks_position('left')

    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    plt.show()
