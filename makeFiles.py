import os
import pandas as pd
import numpy as np
import xlsxwriter
from datetime import date, timedelta


def makeDataFrame(length):
    v = np.random.normal(loc = 10, scale = 4, size = 100)
    soc = np.ceil(v)
    soc[soc < 0] = 0
    headDepartments = ["Chief Cat Office", "Chief Chief Office", "Chief Excel Office", "Chief Chocolate Office"]
    
    df = pd.DataFrame(data = {
        "SoC": soc,
        "headDept": np.random.choice(headDepartments, length)
    })
    return(df)


def mkdirSave(path):
    try:
        os.mkdir(path)
    except: 
        print("Making of folder failed, does it possibly already exist?")
    else:
        print("Folder created")


def writeToExcelFile(df, path, overwrite = False, writeIndex = False):
    if(not overwrite):
        if(os.path.isfile(path)):
            raise OSError("File already exists: %s " % path)
    try:
        with pd.ExcelWriter(
            path = path 
            ) as writer:
            df.to_excel(writer, index = writeIndex) 
    except Exception as e:
        print("Excel file could not be writen, maybe it already exsits?")
        logging.error(traceback.format_exc())
    

def main():
    path = "aLotOfExcelFiles"
    print("Making folder: %s in the current path %s" % (path, os.getcwd()))
    mkdirSave(path)
    
    NOF = 200
    today = date.today()
    for i in range(NOF):
        df = makeDataFrame(100)
        d = today - timedelta(days = i)
        writeToExcelFile(
                df = df,
                path = os.path.join(path, "soc_" + d.strftime("%Y_%m_%d") + ".xlsx")
        )
        

if __name__ == "__main__":

    main()




