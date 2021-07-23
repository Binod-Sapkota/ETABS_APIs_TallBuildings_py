# mender stress-strain model for confined rectangular section
# By Binod Sapkota
# 12/18/2019

import numpy as np
import matplotlib.pyplot as plt


def mander_unconfined(f_c):
    eco = 0.002  # unconfined concrete strain (generally assumed 0.002)
    eu = 0.003
    ec_range = np.arange(0.00001, eu, 0.00005)
    Ec = 5000 * np.sqrt(f_c)
    Esec = f_c / eco
    r = Ec / (Ec - Esec)

    fc = [(f_c *ec / eco* r / (r - 1 + (ec / eco) ** r)) if ec <= 2*eco else
          (2 * f_c * r / (r - 1 + 2 ** r) *((eu-ec)/(eu-2*eco))) for ec in ec_range]

    return ec_range, fc


def mander_circular(ds, fyh, f_c, b, d, c):
    eco = 0.002  # unconfined concrete strain (generally assumed 0.002)
                           # ds = 12   # diameter of hoop
    nsx = 4  # number of leg in x
    nsy = 4  # number of leg in y

    sh = 100  # vertical spacing of hoop

    nlongBarDia = 16
    longBarDia = 23.85
    As_long = nlongBarDia * np.pi * longBarDia ** 2 / 4

    Asx = nsx * np.pi * ds ** 2 / 4  # Area of hoop in x
    Asy = nsy * np.pi * ds ** 2 / 4  # Area of hoop in y

    cb = c  # clear cover to longitudinal bar
    cd = c  # clear cover to longitudinal bar

    cb_hoop = cb - ds / 2
    cd_hoop = cd - ds / 2

    # wi spacing between longitudinal longBarDias
    wi = np.array([90, 90, 90, 90, 90, 90, 90, 90])  # only considered for two adjacent face so double the area later
    w = wi ** 2 / 6  # unconfined area inside hoop

    Aw = 2 * np.sum(w)
    s = sh - ds  # clear spacing between hoops

    bc = b - 2 * cb_hoop  # cb_hoop = cover to hoop
    dc = d - 2 * cd_hoop

    psx = Asx / (bc * sh)
    psy = Asy / (dc * sh)

    Ae = (bc * dc - Aw) * (1 - s / (2 * bc)) * (1 - s / (2 * dc))
    Acc = bc * dc - As_long

    f1x = fyh * psx
    f1y = fyh * psy
    ke = Ae / Acc
    f_1x = ke * f1x/f_c
    f_1y = ke * f1y/f_c
    print(f_1x,f_1y)

    f_1 = ke * f1x
    f_cc = f_c * (2.254 * np.sqrt(1 + 7.94 * f_1 / f_c) - 2 * f_1 / f_c - 1.254)

    ecc = eco * (1 + 5 * (f_cc / f_c - 1))
    ec = np.arange(0.00001, 0.05, 0.00005)
    x = ec / ecc
    Ec = 5000 * np.sqrt(f_c)
    Esec = f_cc / ecc
    r = Ec / (Ec - Esec)

    fc = f_cc * x * r / (r - 1 + x ** r)

    return ec, fc


def Mander_Rectangle(ds, fyh, f_c, b, d, c):
    eco = 0.002  # unconfined concrete strain (generally assumed 0.002)
    # ds = 12   # diameter of hoop
    nsx = 4  # number of leg in x
    nsy = 4  # number of leg in y

    sh = 100  # vertical spacing of hoop

    nlongBarDia = 16
    longBarDia = 23.85
    As_long = nlongBarDia * np.pi * longBarDia ** 2 / 4

    Asx = nsx * np.pi * ds ** 2 / 4  # Area of hoop in x
    Asy = nsy * np.pi * ds ** 2 / 4  # Area of hoop in y

    cb = c  # clear cover to longitudinal bar
    cd = c  # clear cover to longitudinal bar

    cb_hoop = cb - ds / 2
    cd_hoop = cd - ds / 2

    # wi spacing between longitudinal longBarDias
    wi = np.array([108.5, 108.5, 108.5, 108.5, 108.5, 108.5, 108.5, 108.5])  # only considered for two adjacent face so double the area later
    w = wi ** 2 / 6  # unconfined area inside hoop

    Aw = 2 * np.sum(w)
    s = sh - ds  # clear spacing between hoops

    bc = b - 2 * cb_hoop  # cb_hoop = cover to hoop
    dc = d - 2 * cd_hoop

    psx = Asx / (bc * sh)
    psy = Asy / (dc * sh)

    Ae = (bc * dc - Aw) * (1 - s / (2 * bc)) * (1 - s / (2 * dc))
    Acc = bc * dc - As_long

    f1x = fyh * psx
    f1y = fyh * psy
    ke = Ae / Acc
    f_1x = ke * f1x/f_c
    f_1y = ke * f1y/f_c
    print(f_1x,f_1y)

    f_1 = ke * f1x
    f_cc = f_c * (2.254 * np.sqrt(1 + 7.94 * f_1 / f_c) - 2 * f_1 / f_c - 1.254) # this formula is not correct. chart
                                                                                 # should be used

    ecc = eco * (1 + 5 * (f_cc / f_c - 1))
    ec = np.arange(0.00001, 0.05, 0.00005)
    x = ec / ecc
    Ec = 5000 * np.sqrt(f_c)
    Esec = f_cc / ecc
    r = Ec / (Ec - Esec)

    fc = f_cc * x * r / (r - 1 + x ** r)

    return ec, fc


def plotStressStrain():
    ds = [8, 8]
    fyh = [415, 415]
    f_c = 39
    b = 650
    d = 650
    c = 40
    fig = plt.figure(figsize=(8, 5))
    ax = fig.subplots()

    ax.grid(True)

    for i in range(len(ds)):
        ec, fc = Mander_Rectangle(ds[0], fyh[i], f_c, b, d, c)
        ax.plot(ec, fc)

    ec_unconfined, fc_unconfined = mander_unconfined(f_c)
    ax.plot(ec_unconfined,fc_unconfined)
    plt.title("Mander Rectangular Stress-Strain Model")
    plt.show()


plotStressStrain()
