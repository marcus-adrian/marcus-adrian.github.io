import yaml
import openpyxl

#Retrieves data from the YAML file and returns a dictionary with the vals.
def getConfigVariables():
	with open('config.yaml') as f:
		data = yaml.load(f, Loader=yaml.FullLoader)
		return data

def copyWorksheet(workbook, sheetName):
	source = workbook.active
	target = workbook.copy_worksheet(source)
	target.title = sheetName
	return target

def findRowWithKey(worksheet, key):
	rowNum = 0 #loop through and find row for key
	for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=1, values_only=True):
		rowNum += 1
		for cell in row:
			if(cell != None and isinstance(cell, str)):
				if(key in cell):
				   return rowNum
	return 0

def removeRowsBelow(worksheet):
	rowNum = 0 #loop through and find row for mass flows
	massFlowArr = []
	for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=1, values_only=True):
		rowNum += 1
		for cell in row:
			if(cell == "Mass Flows"):
			   massFlowArr.append(rowNum) #More than one intance of mass flow in sheet
	worksheet.delete_rows(massFlowArr[0] + 1, worksheet.max_row) #Delete from first mass flow cell down

def removeZeroRows(worksheet):
    maxRow = worksheet.max_row
    maxCol = worksheet.max_column
    rowNum = 3
    
    for row in worksheet.iter_rows(min_col=3, max_col=maxCol, min_row=3, max_row=maxRow, values_only=True):
        allString = False
        zeroFlag = False #Assume all zero
        for item in row:
            if(isinstance(item, (float, int))): #Is a number?
                if(item != None):
                    if(item != 0):
                        zeroFlag = True
        for item in row:
            if (isinstance(item, str)):
                allString = True
        if(zeroFlag == False and allString == False):
            worksheet.delete_rows(rowNum, 1)

        rowNum += 1
#Step 6
def initialCleanup(worksheet):
	#Lesson, do not use array for deleting rows because they change dynamically per each deletion
	worksheet.delete_rows(1, 2)
	worksheet.delete_rows(findRowWithKey(worksheet, "Maximum Relative Error"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Cost Flow"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "MIXED Substream"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Mass Vapor Fraction"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Mass Liquid Fraction"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Mass Solid Fraction"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Mass Enthalpy"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Mass Entropy"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Mass Density"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Molar Liquid Fraction"), 1)
	worksheet.delete_rows(findRowWithKey(worksheet, "Molar Solid Fraction"), 1)
	removeRowsBelow(worksheet)
	removeZeroRows(worksheet)
	enthalpy_flow = findRowWithKey(worksheet, "Enthalpy Flow")
	enthalpy_flow_units = "B" + str(enthalpy_flow)
	if(worksheet[enthalpy_flow_units].value == "W"):
		print("Enthalpy Flow in Watts, converting to MW")
		for col in worksheet.iter_cols(min_col=3, max_col=worksheet.max_column, 
				min_row=enthalpy_flow, max_row=enthalpy_flow):
			if(col[0].value != None):
				newVal = col[0].value / 1000000
				col[0].value = newVal
def main():
	print("Main Source file called")
	inputData = getConfigVariables()
	streamWorkbook = inputData["streamBookName"]
	print("Working on: " + str(streamWorkbook))
	wb = openpyxl.load_workbook(streamWorkbook)
	modifiedWS = copyWorksheet(wb, "Aspen Data Tables Modified")
	modifiedWS_title = inputData["streamTitle"]
	#Begin work on streams workbook
	initialCleanup(modifiedWS)    
	#setupWB(overallWS, overallTitle)
	#entropyCalculations(overallWS)
	wb.save(streamWorkbook)
	print("Completed first workbook - Steps 6-8")

if __name__ == '__main__':
	main()