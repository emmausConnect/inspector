' get number of rows used in sheet 0..n-1
Function usedRows(sheet, col)
    usedRows = UBound(sheet)
END FUNCTION

' get number of columns used in sheet 0..n-1
Function usedCols(sheet, line)
    usedCols = UBound(sheet, 2)
END FUNCTION

' Autofit all cols in the sheet
Function sheetAutoFit(sheet)
	' nothing to do
End Function

' Close the current instance of excel
Function sheetClose(o)
	' nothing to do
End Function

' Get filename with extension compatible with this lib
Function getOutputFile(fname)
    getOutputFile = getCompatOutputFmt(fname, ".csv")
End Function

' Get preferred extension for this lib
Function getPreferredExtension()
	getPreferredExtension = ".csv"
End Function

' Get avaliable extension type
Function getAvaliableExtensions() 
	Dim exts(0)
	exts(0) = ".csv"
	getAvaliableExtensions = exts
End Function

' Create initial sheet of reports
Function sheetCreateInital()
    ReDim sheet(0, 0)
    Set props = CreateObject("Scripting.Dictionary")
    csvAddValueForRange sheet, "A1", "SUIVI"
    csvAddValueForRange sheet, "E1", "MATERIEL"
    csvAddValueForRange sheet, "J1", "DON"
    csvAddValueForRange sheet, "P1", "CATEGORISATION ET CALCUL DU PRIX DE VENTE"
    csvAddValueForRange sheet, "AB1", "SUIVI DU RECONDITIONNEMENT"
    csvAddValueForRange sheet, "AI1", "VENTE"
    csvAddValueForRange sheet, "AL1", "FICHE TECHNIQUE"

    sheetCreateRow sheet, 2, getTitlesMap()

    props.Add "sheet", sheet
    Set sheetCreateInital = props
End Function


' Dump content of the sheet
' 0..n-1 line from which to begin
Function ggetTab(sheet, line)
    Dim x, y, res
    res = ""
    y = line
    Do While y<UBound(sheet,1)
        x = 0
        Do While x<=UBound(sheet,2)
            res = res & sheet(y,x) & " "
            x=x+1
        Loop
        res = res & "|" & vbCrLf
        y=y+1
    Loop
    ggetTab = res
End Function

' create a sheet row for the provided in a hashmap
' line is 1..n based
' returns the new sheet with new values
Function sheetCreateRowFromHashMap(sheet, line, map)
	FOR EACH k IN map.Keys
		csvAddValueForRange sheet, k&line, map(k)
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

	Dim y, res, value, position
	res = -1
	y = 1
	Do While y<=UBound(sheet)
		position = positions("no_serie") & y
		value = csvGetValueForRange(sheet, position)
		if value=serialNumber then
			res = y
			Exit Do
		end if
		y=y+1
	Loop
	sheetThisPCinSheet = res
End Function
	
' line 0..n-1 line from which to begin
Function ggetCsv(sheet, line)
    Dim x, y, res
    res = ""
    y = line
    Do While y<UBound(sheet,1)
        x = 0
        Do While x<=UBound(sheet,2)
            res = res & """" & sheet(y,x)
            x=x+1
	    res = res & """"
            if x<=UBound(sheet, 2) then
                res = res & ","
            end if
        Loop
        res = res & vbCrLf
        y=y+1
    Loop
    ggetCsv=res
End Function
	
Function getTab(sheet) 
	getTab = ggetTab(sheet,0)
End Function

' Dump content of the sheet
Function tabDump(sheet)
    MsgBox(getTab(sheet))
End Function

Function csvDump(sheet)
    MsgBox(getCsv(sheet))
End Function

Function getCsv(sheet)
    getCsv=ggetCsv(sheet,0)
End Function

' get the array startCol endCol startRow endRow, they are all based 1..n
Function rangeToArray(range)
    Dim startRow, endRow, startCol, endCol, arr
    arr = Split(range, ":")
    if UBound(arr)=0 then
        arr = Array(range, range)
    end if

    ' Gets column, row based on 1..n position
    startCol = stringToNumberValue(reReplace(arr(0),"\d+",""))
    endCol = stringToNumberValue(reReplace(arr(1),"\d+",""))
    startRow = CInt(reReplace(arr(0),"[a-zA-Z]+",""))
    endRow = CInt(reReplace(arr(1),"[a-zA-Z]+",""))
    rangeToArray = Array(startCol,endCol,startRow,endRow)
End Function

' Create a row in the given by getPositionsIndex
' line is 1..n based
Function sheetCreateRowFromArray(sheet, line, data)
    Dim keys, cell, cellv
    keys = positions.Keys()
    for i=0 to UBound(data)
	if i>UBound(keys) then
		Exit For
	End If
        cell = positions(keys(i)) & line
        cellv = data(i)
        csvAddValueForRange sheet, cell, cellv
    next
End Function



' currently format avaliable: A1:A1, A1
' due to tmp assignation you must ensure that you get the return value of this function back
' after the call.
Function csvAddValueForRange(sheet, range, value)
    Dim startRow, endRow, startCol, endCol, arr, x, y, tmp

    value = reReplace(value, vbCrLf, "")

    ' Gets column, row based on 1..n position
    arr = rangeToArray(range)
    startCol = arr(0)
    endCol = arr(1)
    startRow = arr(2)
    endRow = arr(3)

    ' x, y are base 0..n-1
    y = startRow - 1
    'MsgBox("on y: redim from (" & UBound(sheet, 1) & ", " & UBound(sheet,2) & ") to (" & endRow & "," & UBound(sheet,2) & ")")
    ReDimPreserve sheet, endRow, UBound(sheet,2)
    'MsgBox("adding " & value & " to " & range & ",startCol=" & startCol & ",endCol=" & endCol & ",startRow=" & startRow & ",endRow=" & endRow )
    Do While y<endRow
        x = startCol - 1

        'MsgBox("on x: redim from (" & UBound(sheet, 1) & ", " & UBound(sheet,2) & ") to (" & UBound(sheet,1) & "," & endCol & ")")
        ReDimPreserve sheet, UBound(sheet,1), endCol
        Do While x<endCol
	    'MsgBox("add value " & value & " to (" & y & "," & x & ")")
            sheet(y, x) = value
            x=x+1
        Loop
        y=y+1
    Loop
    
    'csvDump(sheet)
End Function

' Get value for the cell specifier given in entry
Function csvGetValueForRange(sheet, range)
    Dim startRow, endRow, startCol, endCol, arr

    arr = rangeToArray(range)
    startCol = arr(0)
    endCol = arr(1)
    startRow = arr(2)
    endRow = arr(3)

    csvGetValueForRange = sheet(startRow - 1, startCol - 1)
End Function

' returns the column index for a given string 0..n-1
' returned column is 1..n
Function stringToNumberValue(s)
    Dim i, char, n, ring
    stringToNumberValue = 0
    i=1
    do while i<=Len(s)
        char = Mid(s, i, 1)
        n = charToNumberValue(char)
        ring = 26 ^ (Len(s) - i)
        stringToNumberValue = stringToNumberValue + ring * (n+1)
        i=i+1
    loop
    stringToNumberValue = stringToNumberValue
End Function


' Write the sheet to the storage
Function sheetWrite(o, f)
	Dim sheet
	sheet = o("sheet")
	f = getAbsoluteFilenameFromRelative(f)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.OpenTextFile(f, 2, True)
	file.Write(getCsv(sheet))
End Function

' Open an existing sheet
Function openExisting(fname)
	Dim line, arr, x, y
	ReDim sheet(0, 0)
	Set props = CreateObject("Scripting.Dictionary")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.OpenTextFile(fname, 1)
	y = 0
	Do While file.AtEndOfStream=False
		line = file.ReadLine
		arr = Split(line, ",")
		x = 0
		Do while x<UBound(arr)
			arr(x)=Mid(arr(x),2,Len(arr(x))-2)
			x=x+1
		Loop
		sheetCreateRow sheet, (file.Line-1), arr
		y = y + 1
	Loop
	props.Add "sheet", sheet
	Set openExisting = props
End Function

'
' Tests
'
ReDim tests(1)
tests(0) = False

IF tests(0) THEN

    ' sheetCreateInital
    Set o = sheetCreateInital()
    assert UBound(o("sheet"), 1)>0, "ubound 1 of sheet"
    assert UBound(o("sheet"), 2)>1, "ubound 2 of sheet"

    ' ReDimPreserve
    ReDim aaa(1,1)
    aaa(0,0)=1
    assert UBound(aaa,1)=1, "UBound(aaa,1)=1"
    assert UBound(aaa,2)=1, "UBound(aaa,2)=1"
    ReDimPreserve aaa, 3, 3
    assert UBound(aaa,1)=3, "UBound(aaa,1)=3"
    assert UBound(aaa,2)=3, "UBound(aaa,2)=3"
    assert aaa(0,0)=1, "aaa(0,0)=1"

    ' IsUpper
    assert IsUpper("aa")=False, "1.IsUpper"
    assert IsUpper("A"), "2.IsUpper"

    ' reReplace
    Dim abb
    abb = "abb"
    abb=reReplace(abb,"bb","bc")
    assert abb="abc", "abc fail"


    ' rangeToArray
    Dim aaar
    aaar = rangeToArray("A2")
    assert aaar(0)=1 and aaar(1)=1, "0111"
    assert aaar(2)=2 and aaar(3)=2, "2131"
    aaar = rangeToArray("AA1")
    assert aaar(0)=27 and aaar(1)=27, "10111"
    assert aaar(2)=1 and aaar(3)=1, "12131"
    aaar = rangeToArray("AB1")
    assert aaar(0)=28 and aaar(1)=28, "20111"
    assert aaar(2)=1 and aaar(3)=1, "22131"
    aaar = rangeToArray("BA1")
    assert aaar(0)=53 and aaar(1)=53, "30111"
    assert aaar(2)=1 and aaar(3)=1, "32131"

    ' stringToNumberValue
    assert stringToNumberValue("BA")=53, "53"
    assert stringToNumberValue("A")=1, "1"

END IF
