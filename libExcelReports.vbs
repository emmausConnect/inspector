'https://docs.microsoft.com/fr-fr/office/vba/api/excel.xlfileformat
Set excelFmts = CreateObject("Scripting.Dictionary")
excelFmts.Add "xla", 18    	'Macro complémentaire Microsoft Excel 97-2003	*.xla
excelFmts.Add "txt", -4158 	'Texte la plateforme actuelle	*.txt
excelFmts.Add "xlsb", 50   	'Classeur Excel binaire	*.xlsb
excelFmts.Add "xls", -4143    	'Classeur normal	*.xls
excelFmts.Add "html", 44
excelFmts.Add "htm", 44
excelFmts.Add "ods", 60
excelFmts.Add "xlsx", 51	'Classeur Open XML	*.xlsx
excelFmts.Add "xltx", 54	'Modèle Open XML	*.xltx
excelFmts.Add "xltm", 53	'Modèle Open XML avec macros	*.xltm
excelFmts.Add "xlsm", 52	'Classeur Open XML avec macros	*.xlsm
excelFmts.Add "xlt", 17		'Format de modèle Excel	*.xlt
excelFmts.Add "mht", 45 	'Archive Web	*.mht; *.mhtml
excelFmts.Add "mhtml", 45	'Archive Web	*.mht; *.mhtml
excelFmts.Add "xml", 46		'Feuille de calcul XML	*.xml
''' Formats fails
'' saveas method fail
'excelFmts.Add "wj2", 14		'Japonais 1-2-3	*.wj2
'excelFmts.Add "wj3", 41		'Japonais 1-2-3	*.wj3
'excelFmts.Add "wk1", 31		'Format Lotus 1-2-3	*.wk1
'excelFmts.Add "wk3", 32 	'Format Lotus 1-2-3	*.wk3
'excelFmts.Add "wk4", 38		'Format Lotus 1-2-3	*.wk4
'excelFmts.Add "wks", 4		'Format Lotus 1-2-3	*.wks
'excelFmts.Add "wq1", 34		'Format Quattro Pro	*.wq1
'excelFmts.Add "csv", 62    	'CSV	 *.csv UTF8 csv
'excelFmts.Add "dbf", 11    	'Format Dbase 4	*.dbf
'' object requis sheet
'excelFmts.Add "xlam", 55 	'Macro complémentaire Open XML	*.xlam
''' Not well supported
'excelFmts.Add "slk", 2		'Format SYLK (Symbolic Link Format)	*.slk
'excelFmts.Add "dif", 9	   	'Format DIF (Data Interchange Format)	*.dif

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
    Set titles = convertBigTitlesToLongRgb(getBigTitles())
    Set props = CreateObject("Scripting.Dictionary")
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False

    Set w = objExcel.Workbooks.Add()
    w.Activate
    With w
     .Title = getSheetTitle()
     .Subject = getSheetSubject()
     .Author = getSheetAuthor()
    End With

    Set sheet = w.ActiveSheet
    Set r = sheet.Range("A1:D1")
    r.Merge
    r.Value = titles("suivi")("text")
    r.Interior.Color = titles("suivi")("bg")
    r.Font.Color = titles("suivi")("text.color")
    Set r = sheet.Range("E1:I1")
    r.Merge
    r.Value = titles("material")("text")
    r.Interior.Color = titles("material")("bg")
    r.Font.Color = titles("material")("text.color")
    Set r = sheet.Range("J1:O1")
    r.Merge
    r.Value = titles("don")("text")
    r.Interior.Color = titles("don")("bg")
    r.Font.Color = titles("don")("text.color")
    Set r = sheet.Range("P1:AA1")
    r.Merge
    r.Value = titles("cat")("text")
    r.Interior.Color = titles("cat")("bg")
    r.Font.Color = titles("cat")("text.color")
    Set r = sheet.Range("AB1:AH1")
    r.Merge
    r.Value = titles("suivi_recon")("text")
    r.Interior.Color = titles("suivi_recon")("bg")
    r.Font.Color = titles("suivi_recon")("text.color")
    Set r = sheet.Range("AI1:AK1")
    r.Merge
    r.Value = titles("vente")("text")
    r.Interior.Color = titles("vente")("bg")
    r.Font.Color = titles("vente")("text.color")
    Set r = sheet.Range("AL1:AV1")
    r.Merge
    r.Value = titles("teck")("text")
    r.Interior.Color = titles("teck")("bg")
    r.Font.Color = titles("teck")("text.color")

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


Dim xltoleft, xlup
xltoleft = -4159  
xlup = -4162 

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


Dim xlVAlignBottom, xlVAlignCenter, xlVAlignDistributed, xlVAlignJustify, xlVAlignTop, xlHAlignCenter, xlHAlignCenterAcrossSelection, xlHAlignDistributed, xlHAlignFill, xlHAlignGeneral, xlHAlignJustify, xlHAlignLeft, xlHAlignRight
xlVAlignBottom = -4107
xlVAlignCenter = -4108
xlVAlignDistributed = -4117
xlVAlignJustify = -4130
xlVAlignTop = -4160
xlHAlignCenter = -4108
xlHAlignCenterAcrossSelection = 7
xlHAlignDistributed = -4117
xlHAlignFill = 5
xlHAlignGeneral = 1
xlHAlignJustify = -4130
xlHAlignLeft = -4131
xlHAlignRight = -4152

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
           o("w").SaveAs f, excelFmts(onlyExtName(f))
        ELSE 
	   o("w").Save
	END IF
End Function

' Close the current instance of excel
Function sheetClose(o)
	o("objExcel").Quit
End Function

' Get preferred extension for this lib
Function getPreferredExtension()
	getPreferredExtension = "xlsx"
End Function

' Get avaliable extension type
Function getAvaliableExtensions() 
	ReDim exts(0)
	exts(0) = "xlsx"
	FOR EACH k IN excelFmts
		if not k=exts(0) then
			ReDim Preserve exts(UBound(exts)+1)
			exts(UBound(exts)) = k
		end if
	NEXT
	getAvaliableExtensions = exts
End Function
