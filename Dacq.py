import argparse
import xlsDataGet

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-f",type=str,help="the xls file you need read")
    parser.add_argument("--sheet",type=str,help="the sheetname in workbook you choose")
    parser.add_argument("--row", type=int, help="the row data you want know")
    parser.add_argument("--col", type=int, help="the col data you want get")
    parser.add_argument("-o",type=str,default="result.txt",help="file path for exporting col data")

    args = parser.parse_args()
    if args.f is None:
        print("Error: you must input the xls file path")
        exit()
    get = xlsDataGet.GetData(args.f)
    if args.sheet is None:
        print("Error: you must input sheet name you choose.there are sheet names you can choose：\n")
        get.printSheets()
        exit()
    get.chooseSheet(args.sheet)

    if args.row is not None:
        get.printTheRowValue(args.row)
    else:
        if args.col is None:
            print("you must input col number")
            exit()
        get.ColValueWrite(args.col, args.o)