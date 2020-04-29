import yaml
import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

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

def addTitle(worksheet, title):
	worksheet.insert_rows(1)
	worksheet['A1'] = title

def addInOutRows(worksheet):
	worksheet.insert_rows(2)
	worksheet['A2'] = "Section In"
	worksheet.insert_rows(3)
	worksheet['A3'] = "Section Out"
	worksheet.merge_cells('A2:CO2') #WHAT A CRAZY FIX WHY
	worksheet.unmerge_cells('A2:CO2') #Three hours of my life gone

def entropyCalculations(worksheet):
	rowIdx = 0
	maxRow = worksheet.max_row
	maxCol = worksheet.max_column
	flag = False
	for row in worksheet.iter_rows(min_row=1, min_col=1,
		max_col=1, max_row=maxRow, values_only=True):
		rowIdx += 1
		for item in row:
			if(item == "Enthalpy Flow"):
				flag = True #found the one row I needed
				newRow = rowIdx + 1
				worksheet.insert_rows(newRow, 2) #Create new row below Enthalpy
				titleCell = "A" + str(newRow)
				unitCell = "B" + str(newRow)
				worksheet[titleCell] = "Entropy Flow" #Add new name for row
				worksheet[unitCell] = "kW/K"
				titleCell = "A" + str(newRow + 1)
				unitCell = "B" + str(newRow + 1)
				worksheet[titleCell] = "Exergy Flow" #Add new name for row
				worksheet[unitCell] = "MW"
		if(flag):
			break

	conversionFactor = 3600 * 1000

	stream_molar_flow = findRowWithKey(worksheet, "Mole Flows")
	stream_molar_flow_units = "B" + str(stream_molar_flow)
	stream_molar_entropy = findRowWithKey(worksheet, "Molar Entropy")
	stream_molar_entropy_units = "B" + str(stream_molar_entropy)

	for row_cells in worksheet.iter_rows(min_row=newRow, max_row=newRow,
		min_col=3, max_col=maxCol):
		for cell in row_cells:
			firstValue = worksheet[str(cell.column_letter) + str(stream_molar_entropy)]
			secondValue = worksheet[str(cell.column_letter) + str(stream_molar_flow)]
			if(firstValue.value != None and secondValue.value != None):
				if(worksheet[stream_molar_flow_units].value == "kmol/hr" and 
					worksheet[stream_molar_entropy_units].value == "J/kmol-K"):
					calculatedValue = (firstValue.value * secondValue.value)/conversionFactor
				else:
					calculatedValue = 1
					print("Unit error in calculating Entropy Flow, required: \n")
					print("Mole Flow Rate: kmol/hr")
					print("Entropy Mixture: J/kmol-K")
				cell.value = calculatedValue
			#Exergy Calculations
			firstValue = worksheet[str(cell.column_letter) + str(findRowWithKey(worksheet, "Enthalpy Flow"))]
			secondValue = worksheet[str(cell.column_letter) + str(findRowWithKey(worksheet, "Entropy Flow"))]
			if(firstValue.value != None and secondValue.value != None):
				calculatedValue = firstValue.value - (0.3 * secondValue.value)
				cell.offset(row=1).value = calculatedValue

	worksheet.delete_rows(findRowWithKey(worksheet, "Description"), 1)

#Returns column letter of first blank after to and from rows
def findBlank(worksheet):
	to_idx = findRowWithKey(worksheet, "To")
	from_idx = findRowWithKey(worksheet, "From")

	for col in worksheet.iter_cols(min_row=to_idx, max_row=to_idx, min_col=3, max_col=worksheet.max_column):
		for cell in col:
			to_cell = str(cell.column_letter) + str(to_idx)
			from_cell = str(cell.column_letter) + str(from_idx)
			if worksheet[to_cell].value == None and worksheet[from_cell].value == None:
				return cell.column_letter
				#worksheet.delete_cols(cell.column_letter,1)

def removeColumns(worksheet):
	to_idx = findRowWithKey(worksheet, "To")
	from_idx = findRowWithKey(worksheet, "From")
	endCol = column_index_from_string(findBlank(worksheet))
	delArray = []
	for col in worksheet.iter_cols(min_row=to_idx, max_row=to_idx, min_col=3, max_col=endCol):
		for cell in col:
			to_cell = str(cell.column_letter) + str(to_idx)
			from_cell = str(cell.column_letter) + str(from_idx)
			if worksheet[to_cell].value != None and worksheet[from_cell].value != None:
				delArray.append(column_index_from_string(cell.column_letter))
	for i in reversed(delArray):
		worksheet.delete_cols(i,1)

def addInOutValues(worksheet):
	to_idx = findRowWithKey(worksheet, "To")
	from_idx = findRowWithKey(worksheet, "From")
	endCol = column_index_from_string(findBlank(worksheet))
	for col in worksheet.iter_cols(min_row=to_idx, max_row=to_idx, min_col=3, max_col=endCol):
		for cell in col:
			write_cell = str(cell.column_letter) + str(1)
			to_cell = str(cell.column_letter) + str(to_idx)
			from_cell = str(cell.column_letter) + str(from_idx)
			if worksheet[to_cell].value != None:
				worksheet[write_cell] = "In"
			if worksheet[from_cell].value != None:
				worksheet[write_cell] = "Out"

def stepSix(worksheet):
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

def stepSeven(worksheet,title):
	addTitle(worksheet,title)
	addInOutRows(worksheet)
	entropyCalculations(worksheet)
	freezeCells = worksheet['C7']
	worksheet.freeze_panes = freezeCells

def stepEight(worksheet):
	freezeCells = worksheet['C7']
	worksheet.freeze_panes = freezeCells
	removeColumns(worksheet)
	addInOutValues(worksheet)


def main():
	print("Main Source file called")
	inputData = getConfigVariables()
	streamWorkbook = inputData["streamBookName"]
	print("Working on: " + str(streamWorkbook))
	wb = openpyxl.load_workbook(streamWorkbook)
	modifiedWS = copyWorksheet(wb, "Aspen Data Tables Modified")
	#Begin work on streams workbook
	stepSix(modifiedWS)
	stepSeven(modifiedWS, inputData["streamTitle"]) 
	wb.save(streamWorkbook)
	overall = wb.copy_worksheet(modifiedWS)
	overall.title = "Overall"
	stepEight(overall)
	wb.save(streamWorkbook)
	print("Completed first workbook - Steps 6-8")

if __name__ == '__main__':
	main()