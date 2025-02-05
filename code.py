import sys, os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.styles.borders import Border, Side

def xlsxToJson(file_name,sheetNumber):
    wb = load_workbook(filename=file_name)
    sheet = wb[wb.sheetnames[sheetNumber-1]]
    head = []
    data = []
    counter = 0
    for fila in sheet:
        #Guardamos los headers
        if counter < 1:
            for cell in fila:
                head.append(cell.value)
            counter += 1
            continue

        #Una vez guardados los headers procedemos a guardar con ellos los datos de cada columna en las filas
        indexOfCell = 0
        cellData = {}
        for cell in fila:
            cellData[head[indexOfCell]] = cell.value
            indexOfCell += 1

        data.append(cellData)
    return data

def generateInstanceOfExcelWithJson(data):
    def getHeaders(data):
        headers = []
        example = data[0]
        for h in example:
            headers.append(h)
        return headers

    def getHeadersLetters(data):
        LETTERS_TO_INDEX = {"0": "A","A": "B","B": "C","C": "D","D": "E","E": "F","F": "G","G": "H","H": "I","I": "J","J": "K","K": "L","L": "M","M": "N","N": "O","O": "P","P": "Q","Q": "R","R": "S","S": "T","T": "U","U": "V","V": "W","W": "X","X": "Y","Y": "Z"}
        headersLetters = {}
        example = data[0]
        actualLetter = "0"
        for h in example:
            headersLetters[h] = LETTERS_TO_INDEX[actualLetter]
            actualLetter = LETTERS_TO_INDEX[actualLetter]
        return headersLetters

    def setHeadData(wb,sheet,value,column,index,color):
        actualSheet = wb[sheet]
        actualSheet[f"{column}{index}"].value = value
        if value == "0":
            actualSheet[f"{column}{index}"].value = ''
        actualSheet[f"{column}{index}"].fill = PatternFill(start_color=color, end_color=color, fill_type = "solid")
        actualSheet[f"{column}{index}"].alignment = Alignment(horizontal='center', vertical='center')
        actualSheet[f"{column}{index}"].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        actualSheet[f"{column}{index}"].font = Font(bold=True, color='ffffff')
    def setData(wb,sheet,value,column,index,color):
        actualSheet = wb[sheet]
        actualSheet[f"{column}{index}"].value = value
        if value == "0":
            actualSheet[f"{column}{index}"].value = ''
        actualSheet[f"{column}{index}"].fill = PatternFill(start_color=color, end_color=color, fill_type = "solid")
        actualSheet[f"{column}{index}"].alignment = Alignment(horizontal='center', vertical='center')
        actualSheet[f"{column}{index}"].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    
    headers = getHeaders(data)
    headersLetters = getHeadersLetters(data)

    wb = Workbook()
    ws = wb.active
    index = 1
    for head in headers:
        ws.column_dimensions[headersLetters[head]].width = 40
        setHeadData(wb,ws.title,head,headersLetters[head],index,'74ac44')

    index = 2

    for row in data:
        if index % 2:
            colorActual = 'e2efda'
        else:
            colorActual = 'ffffff'

        for key in row:
            setData(wb,ws.title,row[key],headersLetters[key],index,colorActual)

        index += 1

	for column in ws.columns:
		max_length = 0
		column_letter = column[0].column_letter
		for cell in column:
		    try:
			if len(str(cell.value)) > max_length:
			    max_length = len(cell.value)
		    except:
			pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    wb['Sheet'].freeze_panes = wb['Sheet']["B2"]
    userProfilePath = os.environ['USERPROFILE']
    tempRoute = f'{userProfilePath}\\AppData\\Local\\Temp\\temp.xlsx'
    wb.save(tempRoute) 
    os.system(f"start EXCEL.EXE {tempRoute}")

generateInstanceOfExcelWithJson(xlsxToJson('temp.xlsx',1))
