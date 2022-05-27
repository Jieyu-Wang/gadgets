import pyautogui as pa
import pandas as pd
import time
import numpy as np

#This program is designed to fill WAF freight matrix with automated search from WorldScale website,
#   with screen in office (Samsung wide screen) and with World Scale website opening on left half of the screen
freight_matrix = pd.read_excel("\\\\rosie\\ldncrude\\WestAfr\\Freight matrix.xlsx",
                               index_col=0)
mileage_matrix = freight_matrix.copy()
load_port_list = list(freight_matrix.index.drop_duplicates())[:]
disport_list = ["Ashkelon"]
previous_data = 0
for disport in disport_list:
    for loadport in load_port_list:

        pa.leftClick(588,371)# click the box for load port
        pa.write(loadport)
        time.sleep(1)
        pa.press("tab")

        time.sleep(1)
        pa.leftClick(1317, 371)# click the box for disport
        pa.write(disport)
        time.sleep(1)
        pa.press("tab")

        pa.leftClick(1439,600)# press search

        time.sleep(3)
        pa.moveTo(x=340,y=476)#,button="left")
        pa.dragTo(1820,477,1.5,button="left")
        pa.hotkey('ctrl', 'c')
        data = pd.io.clipboards.read_clipboard()["LOWEST"][0]#copy and paste to clipboard and read
        mileage = pd.io.clipboards.read_clipboard()["LOWEST"][1]
        try:
            if data == previous_data:
                freight_matrix.loc[loadport,disport] = np.nan
                mileage_matrix.loc[loadport, disport] = np.nan
                print(f"loadport:{loadport} disport:{disport} freight route missing")
            else:
                freight_matrix.loc[loadport,disport] = data
                mileage_matrix.loc[loadport,disport] = mileage
                previous_data=data
                print(f"loadport:{loadport} disport:{disport} freight:{data} mileage:{mileage}")
        except:
            pass

        time.sleep(2)
        pa.leftClick(1727,244)#press new search
        time.sleep(2)
    freight_matrix.to_excel("\\\\rosie\\ldncrude\\WestAfr\\Freight matrix 2022.xlsx")

print("done")
