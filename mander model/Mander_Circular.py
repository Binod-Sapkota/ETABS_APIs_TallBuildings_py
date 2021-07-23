# mender stress-strain model for confined rectangular section
# By Binod Sapkota
# 12/18/2019

import numpy as np
import matplotlib.pyplot as plt
plt.style.use("seaborn-dark")

# ***************** INPUTS **************************************************************
# *************  CROSS-SECTION PROPERTIES ********************
section_dia = 600  # mm
cover = 60.198  # mm (effective cover)
cover = 40
hoopdia = 500 # ds

# *************  REBAR DETAILING ********************
hooptype = 'circular'   # 'spiral'
hoopbardia = 8   # diameter of hoop
hoop_spacing = 100  # vertical spacing of hoop

num_longbar = 16
longbar_dia = 25
longbar_area = num_longbar *np.pi*longbar_dia**2/4

transverse_reinf_area = 100
hoop_clear_spacing = hoop_spacing - hoopbardia

# ************* MATERIAL PROPERTIES ************************
# ...... STEEL .....................
fyh = 413.68  # MPa

# ...... CONCRETE .....................
f_c = 30  # MPa
eco = 0.002  # unconfined concrete strain (generally assumed 0.002)


def mander_circular(ec,f_c):

    core_area = np.pi * hoopdia ** 2 / 4
    pcc = longbar_area / core_area
    if hooptype == 'circular':
        ke = (1-hoop_clear_spacing/(2*hoopdia))**2/(1-pcc)
    else:
        ke = (1-hoop_clear_spacing/(2*hoopdia))/(1-pcc)

    transverse_confinement_volratio = 4*transverse_reinf_area/(hoopdia *hoop_spacing)
    effective_confinement = 0.5*ke*transverse_confinement_volratio *fyh

    f_cc = f_c * (2.254 * np.sqrt(1 + 7.94 * effective_confinement / f_c)
                  - 2 * effective_confinement / f_c - 1.254)

    ecc = eco * (1 + 5 * (f_cc / f_c - 1))
    x = ec / ecc
    Ec = 5000 * np.sqrt(f_c)
    Esec = f_cc / ecc
    r = Ec / (Ec - Esec)
    fc = f_cc * x * r / (r - 1 + x ** r)

    return fc

def plot_mander_circular(f_c):
    ec = np.arange(0.00001, 0.05, 0.00005)
    fig = plt.figure(figsize=(7, 5))
    ax = fig.subplots()

    ax.grid(True)
    fc = mander_circular(ec,f_c)
    ax.plot(ec, fc)
    plt.title("Mander Circular Stress-Strain Model")
    plt.xlabel("Strain")
    plt.ylabel("Compressive Stress, MPa")
    # plt.show()
    return fig

# plot_mander_circular()
