




' Create a row in the given by getPositionsIndex
Function sheetCreateRowFromArray(sheet, line, data) 
	Dim keys, cell, cellv
        keys = positions.Keys()
       for i=0 to UBound(data)
	    cell = positions(keys(i)) & line
	    cellv = data(i)
	    sheet.Range(cell).Value = cellv
       next
End Function

' create a sheet row for the provided in a hashmap
Function sheetCreateRowFromHashMap(sheet, line, map)
	FOR EACH k IN map.Keys
		sheet.Range(k&line).Value = map(k)
	NEXT
End Function

' Create initial sheet of reports
' https://docs.microsoft.com/fr-fr/office/vba/api/excel.application(object)
' https://docs.microsoft.com/fr-fr/office/vba/api/excel.worksheet
Function sheetCreateInital()
    Set props = CreateObject("Scripting.Dictionary")
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False

	Set w = objExcel.Workbooks.Add() 
    w.Activate
    With w
     .Title = "Tous les reconditionnements" 
     .Subject = "Reconditionnements"
     .Author = "Emmaüs"
    End With

    Set sheet = w.ActiveSheet
    Set r = sheet.Range("A1:D1")
    r.Merge
    r.Value = "SUIVI"
    r.Interior.ColorIndex = 48
    r.Font.ColorIndex = 2
    Set r = sheet.Range("E1:I1")
    r.Merge
    r.Value = "MATERIEL"
    r.Interior.ColorIndex = 23
    r.Font.ColorIndex = 2
    Set r = sheet.Range("J1:O1")
    r.Merge
    r.Value = "DON"
    r.Interior.ColorIndex = 10
    r.Font.ColorIndex = 2
    Set r = sheet.Range("P1:AA1")
    r.Merge
    r.Value = "CATEGORISATION ET CALCUL DU PRIX DE VENTE"
    r.Interior.ColorIndex = 45
    r.Font.ColorIndex = 2
    Set r = sheet.Range("AB1:AH1")
    r.Merge
    r.Value = "SUIVI DU RECONDITIONNEMENT"
    r.Interior.ColorIndex = 55
    r.Font.ColorIndex = 2
    Set r = sheet.Range("AI1:AK1")
    r.Merge
    r.Value = "VENTE"
    r.Interior.ColorIndex = 50
    r.Font.ColorIndex = 1
    Set r = sheet.Range("AL1:AV1")
    r.Merge
    r.Value = "FICHE TECHNIQUE"
    r.Interior.ColorIndex = 46
    r.Font.ColorIndex = 1

    sheetCreateRow sheet, 2, getTitlesMap()

	Set r = sheet.Range("A1:AV2")
    r.Font.Bold = True
	
    props.Add "objExcel", objExcel
    props.Add "w", w
    props.Add "sheet", sheet
    props.Add "mustWrite", True
    Set sheetCreateInital = props
End Function

' Open an existing sheet 
' WARNING : absolute path are not allowed
Function openExisting(fname)
	Set props = CreateObject("Scripting.Dictionary")
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = False
	props.Add "objExcel", objExcel
	Set w = objExcel.Workbooks.Open(fname)
	props.Add "w", w
	props.Add "sheet", w.ActiveSheet
        props.Add "mustWrite", False
	Set openExisting = props
End Function



Const xltoleft = -4159  
Const xlup = -4162 

' get number of rows used in sheet
Function usedRows(sheet, col)
	With sheet
	    usedRows = .Cells(.Rows.Count, col).End(xlup).Row
	End With
END FUNCTION

' get number of columns used in sheet
Function usedCols(sheet, line)
	With sheet
	    usedCols = .Cells(line, .Columns.Count).End(xltoleft).Column
	End With
END FUNCTION



Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160
Const xlHAlignCenter = -4108
Const xlHAlignCenterAcrossSelection = 7
Const xlHAlignDistributed = -4117
Const xlHAlignFill = 5
Const xlHAlignGeneral = 1
Const xlHAlignJustify = -4130
Const xlHAlignLeft = -4131
Const xlHAlignRight = -4152

' Autofit all cols in the sheet
Function sheetAutoFit(sheet)
	Set rows = sheet.Rows
        rows.VerticalAlignment = xlVAlignCenter
        rows.HorizontalAlignment = xlHAlignCenter
	FOR EACH v IN positions.Items
		sheet.Columns(v).Autofit
	NEXT
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
		Set range = sheet.Range(positions("no_serie") & i)
		Set c = range.Item(1)
		if c.Value=serialNumber then
			res = i	
			Exit For
		end if
	next
	sheetThisPCinSheet = res
End Function

' Write the sheet to the storage
Function sheetWrite(o, f)
	f = getAbsoluteFilenameFromRelative(f)
        IF o("mustWrite") THEN
           o("w").Close True, f
        ELSE 
	   o("w").Save
	END IF
End Function

' Close the current instance of excel
Function sheetClose(o)
	o("objExcel").Quit
End Function

' Get filename with extension compatible with this lib
Function getOutputFile(fname)
	getOutputFile = getCompatOutputFmt(fname, ".xlsx")
End Function
