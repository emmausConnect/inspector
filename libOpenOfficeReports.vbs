' Create a row in the given by getPositionsIndex
Function sheetCreateRowFromArray(sheet, line, data) 
       Dim keys, cell, cellv
       keys = positions.Keys()
       for i=0 to UBound(data)
	    cell = positions(keys(i)) & line
	    cellv = data(i)
	    sheet.getCellRangeByName(cell).getCellByPosition(0, 0).String = cellv
       next
End Function

' create a sheet row for the provided in a hashmap
Function sheetCreateRowFromHashMap(sheet, line, map)
	FOR EACH k IN map.Keys
		sheet.getCellRangeByName(k&line).getCellByPosition(0, 0).String  = map(k)
	NEXT
End Function

const FontWeightDONTKNOW = 0.000000 'The font weight is not specified/known.
const FontWeightTHIN = 50.000000 'specifies a 50% font weight.
const FontWeightULTRALIGHT = 60.000000 'specifies a 60% font weight.
const FontWeightLIGHT = 75.000000 'specifies a 75% font weight.
const FontWeightSEMILIGHT = 90.000000 'specifies a 90% font weight.
const FontWeightNORMAL = 100.000000 'specifies a normal font weight.
const FontWeightSEMIBOLD = 110.000000 'specifies a 110% font weight.
const FontWeightBOLD = 150.000000 'specifies a 150% font weight.
const FontWeightULTRABOLD = 175.000000 'specifies a 175% font weight.
const FontWeightBLACK = 200.000000 'specifies a 200% font weight.

' Create initial sheet of reports
'https://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/Desktop.html
Function sheetCreateInital()
	Set props = CreateObject("Scripting.Dictionary")
	Set sm = CreateObject("com.sun.star.ServiceManager")
	Set d = sm.CreateInstance("com.sun.star.frame.Desktop")
	Dim arg()
	Set w = d.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
	d.getCurrentFrame().getContainerWindow().setVisible(False) ' todo : better way to do that

	Set sheet = w.CurrentController.getActiveSheet()
	Set r = sheet.getCellRangeByName("A1:D1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = "SUIVI"
	r.CellBackColor = &H7e8187
	r.CharColor = &HFFFFFF
	Set r = sheet.getCellRangeByName("E1:I1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = "MATERIEL"
	r.CellBackColor = &H0055ff
	r.CharColor = &HFFFFFF
	Set r = sheet.getCellRangeByName("J1:O1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = "DON"
	r.CellBackColor = &H005c05
	r.CharColor = &HFFFFFF
	Set r = sheet.getCellRangeByName("P1:AA1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = "CATEGORISATION ET CALCUL DU PRIX DE VENTE"
	r.CellBackColor = &Hdb852e
	r.CharColor = &HFFFFFF
	Set r = sheet.getCellRangeByName("AB1:AH1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = "SUIVI DU RECONDITIONNEMENT"
	r.CellBackColor = &H2a039e
	r.CharColor = &HFFFFFF
	Set r = sheet.getCellRangeByName("AI1:AK1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = "VENTE"
	r.CellBackColor = &H5de381
	r.CharColor = &H000000
	Set r = sheet.getCellRangeByName("AL1:AV1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = "FICHE TECHNIQUE"
	r.CellBackColor = &Hab521b
	r.CharColor = &H000000

	sheetCreateRow sheet, 2, getTitlesMap()

	Set r = sheet.getCellRangeByName("A1:AV2")
	r.CharWeight = FontWeightBOLD
	props.Add "sm", sm
	props.Add "d", d
	props.Add "w", w
	props.Add "sheet", sheet
	props.Add "mustWrite", True
	Set sheetCreateInital = props
End Function

' Open an existing sheet 
Function openExisting(fname)
	Set props = CreateObject("Scripting.Dictionary")
	Set sm = CreateObject("com.sun.star.ServiceManager")
	Set d = sm.CreateInstance("com.sun.star.frame.Desktop")
	Dim arg()
	Set w = d.loadComponentFromURL(ConvertToURL(fname), "_blank", 0, arg)
	d.getCurrentFrame().getContainerWindow().setVisible(False) ' todo : better way to do that
	props.Add "sm", sm
	props.Add "d", d
	props.Add "w", w
	props.Add "sheet", w.CurrentController.getActiveSheet()
	Set openExisting = props
End Function

' get number of columns used in sheet
Function usedCols(sheet, line)
	Set oCursor = o("sheet").createCursor()
	oCursor.gotoEndOfUsedArea(True)
	Set oColumns = oCursor.getColumns()
	usedCols = oColumns.getCount()
END FUNCTION

' get number of rows used in sheet
Function usedRows(sheet, col)
	Set oCursor = o("sheet").createCursor()
	oCursor.gotoEndOfUsedArea(True)
	Set oRows = oCursor.getRows()
	usedRows = oRows.getCount()
End Function

Const CellHoriJustifySTANDARD = 0 'default alignment is used (left for numbers, right for text).  
Const CellHoriJustifyLEFT = 1 'contents are aligned to the left edge of the cell.  
Const CellHoriJustifyCENTER = 2 'contents are horizontally centered.  
Const CellHoriJustifyRIGHT = 3 'contents are aligned to the right edge of the cell.  
Const CellHoriJustifyBLOCK = 4 'contents are justified to the cell width.  
Const CellHoriJustifyREPEAT = 5 'contents are repeated to fill the cell.
Const CellVertJustifySTANDARD = 0 'default alignment is used.  
Const CellVertJustifyTOP = 1 'contents are aligned with the upper edge of the cell.  
Const CellVertJustifyCENTER = 2 'contents are aligned to the vertical middle of the cell.  
Const CellVertJustifyBOTTOM = 3 'contents are aligned to the lower edge of the cell.  

' Autofit all cols in the sheet
Function sheetAutoFit(sheet)
	Set r = sheet.getCellRangeByName("A1:AV1000")
	r.IsTextWrapped = True
	r.HoriJustify = CellHoriJustifyCENTER
	r.VertJustify = CellVertJustifyCENTER
End Function

' Returns -1 if this pc is not in the sheet else 1..n line where the entry has been found
Function sheetThisPCinSheet(sheet)
        dim serialNumber
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
        For Each objItem in colItems
            serialNumber = objItem.SerialNumber
	Next
	Dim res
	res = -1
	Set rows = sheet.Rows
	For i = 1 To usedRows(sheet,positions("cpu"))
		Set range = sheet.getCellRangeByName(positions("no_serie") & i)
		Set c = range.getCellByPosition(0, 0)
		if c.String=serialNumber then
			res = i	
			Exit For
		end if
	next
	sheetThisPCinSheet = res
End Function

' Convert Windows pathnames to url
Function ConvertToURL(sFileName)
	Dim sTmpFile
	If Left(sFileName, 7) = "file://" Then
		ConvertToURL = sFileName
		Exit Function
	End If
	ConvertToURL = "file:///"
	sTmpFile = getAbsoluteFilenameFromRelative(sFileName)
	' replace any "\" by "/"
	sTmpFile = Replace(sTmpFile,"\","/") 
	' replace any "%" by "%25"
	sTmpFile = Replace(sTmpFile,"%","%25") 
	' replace any " " by "%20"
	sTmpFile = Replace(sTmpFile," ","%20")
	ConvertToURL = ConvertToURL & sTmpFile
End Function

' Write the sheet to the storage
Function sheetWrite(o, f)
	o("w").StoreToURL ConvertToURL(f), Array()
End Function

' Close the current instance of excel
Function sheetClose(o)
	o("w").close(true)
End Function

' Get filename with extension compatible with this lib
Function getOutputFile(fname)
	getOutputFile = getCompatOutputFmt(fname, ".ods")
End Function
