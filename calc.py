import xlutil

file = "/Users/macos/Desktop/test.xlsx"
row = xlutil.getRowCount(file,'working total')

for i in range(3,row):

    print("currently processing row is",i)
    print("...........")
    serial = xlutil.readData(file,'working total',i,1)
    print("serial number is",serial)
    quantity = xlutil.readData(file,'working total',i,5)
    print("quantity is:",quantity)
    balance = xlutil.readData(file,'working total',i,6)
    print("balance is",balance)
    total = xlutil.readData(file,'working total',i,7)
    print("total is",total)
    typeofsale = xlutil.readData(file,'working total',i,8)
    print("type of sale is",typeofsale)
    previous_total_stock = xlutil.readData(file,'working total',i-1,7)
    print("previous stock is",previous_total_stock)
    if typeofsale == "IN":
        quantity_check = int(previous_total_stock) + int(quantity)
        print("calculated in value is ",quantity_check)
        print("original value is",total)
        if quantity_check != total:
            print("******** error ********")
            print("this serial numbered item has issue",serial)
            xlutil.writeData(file,'working total',i,12,"item has stock issue")
    elif typeofsale == "OUT":
        quantity_check = int(previous_total_stock) - int(quantity)
        print("calculated out value is",quantity_check)
        print("original value is", total)
        if quantity_check != total:
            print("******** error ********")
            print("******* this serial numbered item has issue: ", serial)
            xlutil.writeData(file, 'working total', i, 12, "item has stock issue")
    else:
        print("check the format of sheet, format issue occured")


















