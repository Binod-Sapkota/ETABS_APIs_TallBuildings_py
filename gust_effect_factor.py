# This program calculates gust response factor as per ASCE 7-16

from math import sqrt,log,exp

def gust_effect_factor(h=183,B=30,L=30,n1=0.2,V =51,beta = 0.01,exposure= "B"):
    # Table 26.11-1 Terrain Exposure Constants (SI Units)
    terrain_exposure_constants = {"B": [7.0, 365.76, 1/7, 0.84, 1/4.0, 0.45, 0.3, 97.54, 1/3.0, 9.14],
                                  "C": [9.5,274.32, 1/9.5, 1.00, 1/6.5, 0.65, 0.20, 152.40, 1/5.0, 4.57],
                                  "D": [11.5, 213.36, 1/11.5, 1.07, 1/9.0, 0.80, 0.15, 198.12, 1/8.0, 2.13]}

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


    z_bar = 0.6 * h         # equivalent height of structure (=0.6*h)
    alpha_bar = terrain_exposure_constants[exposure][4]
    b_bar = terrain_exposure_constants[exposure][5]
    c = terrain_exposure_constants[exposure][6]
    l = terrain_exposure_constants[exposure][7]
    e_bar = terrain_exposure_constants[exposure][8]
    Iz = c * (10 / z_bar) ** (1 / 6)                  # Intensity_of_turbulence
    gQ = 3.4                                # peak_factor_for_backgroundresponse
    gv = 3.4                                 # peak_factor_for_windresponse
    gR = sqrt(2*log(3600*n1) )+ 0.577/(sqrt(2*log(3600*n1)))      # peak_factor_for_resonantresponse
    Lz = l * (z_bar / 10) ** e_bar
    Q = sqrt(1/(1+0.63*((B+h)/Lz)**0.63))                                  # background_response_factor
    Vz = b_bar * (z_bar / 10) ** alpha_bar * V
    N1 = n1*Lz/Vz
    Rn = 7.47*N1/(1 +10.3*N1)**(5/3)


    nus = [4.6*n1*h/Vz,4.6*n1*B/Vz,15.4*n1*L/Vz]

    R_l = [1/nu - 1/(2*nu**2)*(1-exp(-2*nu)) for nu in nus]

    R = sqrt(1/beta*Rn*R_l[0]*R_l[1]*(0.53+0.47*R_l[2]))             # resonant_response_factor


    gust_effect_factor_flexible = 0.925*(1+1.7*Iz*sqrt(gQ*gQ*Q*Q + gR*gR*R*R))/(1+1.7*gv*Iz)

    cache = (z_bar,alpha_bar,b_bar,c,Iz,Vz,Q*Q,Rn,R_l[0],R_l[1],R_l[2],nus,gR)

    return gust_effect_factor_flexible, cache



factor,cache = gust_effect_factor()


[print(ch) for ch in cache]
print(factor)
