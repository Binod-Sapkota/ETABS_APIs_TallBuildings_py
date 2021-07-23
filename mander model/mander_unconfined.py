import numpy as np
import matplotlib.pyplot as plt
plt.style.use("seaborn-dark")

def mander_unconfined(ec, f_c):
    eco = 0.002  # unconfined concrete strain (generally assumed 0.002)
    eu = 0.003
    Ec = 5000 * np.sqrt(f_c)
    Esec = f_c / eco
    r = Ec / (Ec - Esec)

    if ec <= 2 * eco:
        fc = (f_c *ec / eco* r / (r - 1 + (ec / eco) ** r))
    else:
        fc = (2 * f_c * r / (r - 1 + 2 ** r) *((eu-ec)/(eu-2*eco)))

    return fc


def plot_mander_unconfined(f_c=30):
    eu = 0.003
    ec_all = np.arange(0.00001, eu+0.00005, 0.00005)
    fig = plt.figure(figsize=(7, 5))
    ax = fig.subplots()
    ax.grid(True)
    fc = [mander_unconfined(ec,f_c) for ec in ec_all]
    ax.plot(ec_all, fc)
    plt.title("Mander Unconfined Stress-Strain Model")
    plt.xlabel("Strain")
    plt.ylabel("Compressive Stress, MPa")
    # plt.show()
    return fig

# plot_mander_unconfined()
