Set supportedFormats = CreateObject("Scripting.Dictionary")
supportedFormats.Add "xls", "MS Excel 97"					'Microsoft Excel 97-2003
supportedFormats.Add "ods", "Calc8" 						'ODF Spreadsheet
supportedFormats.Add "html", "HTML (StarCalc)"  				'HTML Document (Calc)
supportedFormats.Add "ots", "calc8_template"					'ODF Spreadsheet Template
supportedFormats.Add "fods", "OpenDocument Spreadsheet Flat XML"		'Flat XML ODF Spreadsheet
supportedFormats.Add "uos", "UOF spreadsheet"					'Unified Office Format Spreadsheet
supportedFormats.Add "xlsx", "Calc Office Open XML" 				'Microsoft Excel 2007-2013 XML
supportedFormats.Add "xlt", "MS Excel 97 Vorlage/Template"			'Microsoft Excel 97-2003 Template
supportedFormats.Add "dif", "DIF"						'Data Interchange Format
supportedFormats.Add "dbf", "dBase"						'dBase
supportedFormats.Add "slk", "SYLK"						'SYLK
supportedFormats.Add "csv", "Text - txt - csv (StarCalc)"			'Text CSV
'todo for csv
'args3(2).Name = "FilterOptions"
'args3(2).Value = "44,34,IBMPC_850,1,,0,false,true,false,false,false"
supportedFormats.Add "xlsm", "Calc MS Excel 2007 VBA XML"		'Microsoft Excel 2007-2016 XML (macro enabled)

' Get avaliable extension type
Function getAvaliableExtensions() 
	ReDim exts(0)
	exts(0) = "ods"
	FOR EACH k IN supportedFormats
		if not k=exts(0) then
			ReDim Preserve exts(UBound(exts)+1)
			exts(UBound(exts)) = k
		end if
	NEXT
	getAvaliableExtensions = exts
End Function

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

Dim FontWeightDONTKNOW, FontWeightTHIN, FontWeightULTRALIGHT, FontWeightLIGHT, FontWeightSEMILIGHT, FontWeightNORMAL, FontWeightSEMIBOLD, FontWeightBOLD, FontWeightULTRABOLD, FontWeightBLACK
FontWeightDONTKNOW = 0.000000 'The font weight is not specified/known.
FontWeightTHIN = 50.000000 'specifies a 50% font weight.
FontWeightULTRALIGHT = 60.000000 'specifies a 60% font weight.
FontWeightLIGHT = 75.000000 'specifies a 75% font weight.
FontWeightSEMILIGHT = 90.000000 'specifies a 90% font weight.
FontWeightNORMAL = 100.000000 'specifies a normal font weight.
FontWeightSEMIBOLD = 110.000000 'specifies a 110% font weight.
FontWeightBOLD = 150.000000 'specifies a 150% font weight.
FontWeightULTRABOLD = 175.000000 'specifies a 175% font weight.
FontWeightBLACK = 200.000000 'specifies a 200% font weight.

' Create initial sheet of reports
'https://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/Desktop.html
Function sheetCreateInital()
	Set titles = getBigTitles()
	Set props = CreateObject("Scripting.Dictionary")
	Set sm = CreateObject("com.sun.star.ServiceManager")
	Set d = sm.CreateInstance("com.sun.star.frame.Desktop")
	Set w = d.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array(getBeanProperty("Hidden", True)))

	Set sheet = w.CurrentController.getActiveSheet()
	Set r = sheet.getCellRangeByName("A1:D1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = titles("suivi")("text")
	r.CellBackColor = titles("suivi")("bg")
	r.CharColor = titles("suivi")("text.color")
	Set r = sheet.getCellRangeByName("E1:I1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = titles("material")("text")
	r.CellBackColor = titles("material")("bg")
	r.CharColor = titles("material")("text.color")
	Set r = sheet.getCellRangeByName("J1:O1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = titles("don")("text")
	r.CellBackColor = titles("don")("bg")
	r.CharColor = titles("don")("text.color")
	Set r = sheet.getCellRangeByName("P1:AA1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = titles("cat")("text")
	r.CellBackColor = titles("cat")("bg")
	r.CharColor = titles("cat")("text.color")
	Set r = sheet.getCellRangeByName("AB1:AH1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = titles("suivi_recon")("text")
	r.CellBackColor = titles("suivi_recon")("bg")
	r.CharColor = titles("suivi_recon")("text.color")
	Set r = sheet.getCellRangeByName("AI1:AK1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = titles("vente")("text")
	r.CellBackColor = titles("vente")("bg")
	r.CharColor = titles("vente")("text.color")
	Set r = sheet.getCellRangeByName("AL1:AV1")
	r.merge(True)
	r.getCellByPosition(0, 0).String = titles("teck")("text")
	r.CellBackColor = titles("teck")("bg")
	r.CharColor = titles("teck")("text.color")

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

' Get property in a bean
' https://wiki.openoffice.org/wiki/Opening_a_document
Function getBeanProperty(name, value)
	Set oSM = CreateObject("com.sun.star.ServiceManager")
	Set oPropertyValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	oPropertyValue.Name = name
	oPropertyValue.Value = value
	Set getBeanProperty = oPropertyValue
End Function

' Open an existing sheet 
Function openExisting(fname)
	Set props = CreateObject("Scripting.Dictionary")
	Set sm = CreateObject("com.sun.star.ServiceManager")
	Set d = sm.CreateInstance("com.sun.star.frame.Desktop")	
	Set w = d.loadComponentFromURL(ConvertToURL(fname), "_blank", 0, Array(getBeanProperty("Hidden", True)))
	props.Add "sm", sm
	props.Add "d", d
	props.Add "w", w
	props.Add "sheet", w.CurrentController.getActiveSheet()
	Set openExisting = props
End Function

' get number of columns used in sheet
Function usedCols(sheet, line)
	Set oCursor = sheet.createCursor()
	oCursor.gotoEndOfUsedArea(True)
	Set oColumns = oCursor.getColumns()
	usedCols = oColumns.getCount()
END FUNCTION

' get number of rows used in sheet
Function usedRows(sheet, col)
	Set oCursor = sheet.createCursor()
	oCursor.gotoEndOfUsedArea(True)
	Set oRows = oCursor.getRows()
	usedRows = oRows.getCount()
End Function

Dim CellHoriJustifySTANDARD, CellHoriJustifyLEFT, CellHoriJustifyCENTER, CellHoriJustifyRIGHT, CellHoriJustifyBLOCK, CellHoriJustifyREPEAT, CellVertJustifySTANDARD, CellVertJustifyTOP, CellVertJustifyCENTER, CellVertJustifyBOTTOM
CellHoriJustifySTANDARD = 0 'default alignment is used (left for numbers, right for text).  
CellHoriJustifyLEFT = 1 'contents are aligned to the left edge of the cell.  
CellHoriJustifyCENTER = 2 'contents are horizontally centered.  
CellHoriJustifyRIGHT = 3 'contents are aligned to the right edge of the cell.  
CellHoriJustifyBLOCK = 4 'contents are justified to the cell width.  
CellHoriJustifyREPEAT = 5 'contents are repeated to fill the cell.
CellVertJustifySTANDARD = 0 'default alignment is used.  
CellVertJustifyTOP = 1 'contents are aligned with the upper edge of the cell.  
CellVertJustifyCENTER = 2 'contents are aligned to the vertical middle of the cell.  
CellVertJustifyBOTTOM = 3 'contents are aligned to the lower edge of the cell.  

' Autofit all cols in the sheet
Function sheetAutoFit(sheet)
	Set r = sheet.getCellRangeByName("A1:AV1000")
	r.IsTextWrapped = True
	r.HoriJustify = CellHoriJustifyCENTER
	r.VertJustify = CellVertJustifyCENTER
End Function

' Returns -1 if this pc is not in the sheet else 1..n line where the entry has been found
Function sheetThisPCinSheet(sheet, serialNumber)
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
	Set oSM = CreateObject("com.sun.star.ServiceManager")	
	Set oPropertyValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	oPropertyValue.Name = "FilterName"
	oPropertyValue.Value = supportedFormats(onlyExtName(f))
	o("w").StoreToURL ConvertToURL(f), Array(oPropertyValue)
End Function

' Close the current instance of excel
Function sheetClose(o)
	o("w").close(true)
End Function

' Get preferred extension for this lib
Function getPreferredExtension()
	getPreferredExtension = "ods"
End Function

