import numpy as np
import matplotlib.pyplot as plt
# from scipy.spatial import ConvexHull, convex_hull_plot_2d
# import alphashape
# from descartes import PolygonPatch
import ObtainResults
# from ObtainResults import get_dataFromTable, getStoryColumnInfo, getStoryWallInfo
import pandas as pd
import os

from matplotlib.animation import FuncAnimation


def storywiseplot(plotColumn=None, FUNC='max', lb=None, ub=None,figsize=(6, 8),
                  xlabel="x", ylabel="Story (%)",title="Story Respose Plot"):

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
    fig = plt.figure(figsize=figsize)
    ax = fig.add_subplot()
    linealpha = 0.5

    SMALL_SIZE = 8
    MEDIUM_SIZE = 10
    BIGGER_SIZE = 22

    plt.rc('font', family='Arial',size=BIGGER_SIZE)  # controls default text sizes
    plt.rc('axes', titlesize=BIGGER_SIZE)  # fontsize of the axes title
    plt.rc('axes', labelsize=MEDIUM_SIZE)  # fontsize of the x and y labels
    plt.rc('xtick', labelsize=SMALL_SIZE)  # fontsize of the tick labels
    plt.rc('ytick', labelsize=SMALL_SIZE)  # fontsize of the tick labels
    plt.rc('legend', fontsize=SMALL_SIZE)  # legend fontsize
    # plt.rc('figure',  titleweight = 'bold',titlesize=BIGGER_SIZE)  # fontsize of the figure title

    xmin=0
    xmax=0
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
                    ax.plot(x, y,alpha=linealpha,linewidth=1)

            elif ub is not None:
                xmax = ub
                if np.nanmax(x) < ub:
                    ax.plot(x, y,alpha=linealpha,linewidth=1)

            elif lb is not None:
                xmin = lb
                if np.nanmax(x) > lb:
                    ax.plot(x, y,alpha=linealpha,linewidth=1)
            else:
                ax.plot(x, y,alpha=linealpha,linewidth=1)
        except:
            pass
    ax.spines['left'].set_position('zero')
    ax.spines['right'].set_color('none')
    ax.spines['bottom'].set_position('zero')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_bounds(0, 100)
    ax.spines['bottom'].set_bounds(xmin, xmax)
    ax.xaxis.set_ticks_position('bottom')
    ax.yaxis.set_ticks_position('left')

    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    plt.show()


# storywiseplot(plotColumn=["VibrationServiceabilityIndex"], xlabel="Story Shear (%)",ub=0.001)

# storywiseplot(plotColumn=["DriftX_%", "DriftY_%"], FUNC='max', ub=0.5)
# ObtainResults.podiumLevelPlot()


# Program to find the Neutral Axis Depth
import numpy as np

b = 200
h = 300
dd = 36

d = h - dd
fck = 38
fy = 420  # MPa
Es = 2 * 10 ** 5  # MPa
Moment = 39.2 * 10 ** 3  # Nm
C = 45  # for trial

while True:
    a = 0.85 * C
    # reinf = [no. of bar, bar dia, bar dist from top,fy ]
    reinf = np.array([[2, 12, d, fy]])#, [2, 9, 251.17, 240], [2, 9, 147.33, 240], [2, 9, 43.5, 240]])

    strain = np.array([0.003 * (C - reinf[0, 2]) / C])#, 0.003 * (C - reinf[1, 2]) / C, 0.003 * (C - reinf[2, 2]) / C,
                     #  0.003 * (C - reinf[3, 2]) / C])

    Abar = reinf[:, 0] * np.pi * reinf[:, 1] * reinf[:, 1] / 4
    #    print(Abar)
    #    print(strain)
    tForce = np.zeros(5)
    tForceTot = 0
    cForceTot = 0
    for i in range(1):
        ey = reinf[i, 3] / Es
        if np.abs(strain[i]) <= ey:
            fs = np.abs(strain[i]) * Es
        else:
            fs = reinf[i, 3]
        #        print(fs)

        if C < reinf[i, 2]:
            tForce[i] = -Abar[i] * fs
            tForceTot = -Abar[i] * fs + tForceTot
        elif C == reinf[i, 2]:
            tForce[i] = 0
        else:
            tForce[i] = Abar[i] * (fs - 0.85 * fck)
            cForceTot = Abar[i] * (fs - 0.85 * fck) + cForceTot

    if fck < 27.58:
        B = 0.85
    elif fck > 27.58 and fck < 55.16:
        B = 0.85 - 0.05 * (fck - 27.58) / 6.9
    else:
        B = 0.65

    tForce[i + 1] = 0.85 * fck * B * b * C
    cForceTot = 0.85 * fck * B * b * C + cForceTot

    if (np.abs(np.abs(tForceTot) - cForceTot)) < 0.5:
        break
    else:
        C = C + 0.0001
    # print(C)

print(C)
distCForce = 0.85 * C / 2
Mn = 0

for i in range(1):
    Mn = Mn + tForce[i] * (h / 2 - reinf[i, 2])

Mn = Mn + tForce[i + 1] * (h / 2 - B * C / 2)
print(tForceTot, cForceTot, Mn)