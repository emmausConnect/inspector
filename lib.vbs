'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Library to extract informations from computer
'It is not oriented to diagnose computers but give an overview of it's features
'doc: https://www.activexperts.com/admin/scripts/wmi/
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strComputer
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

' Get a string describing the type of disk used inside this computer
' https://www.tek-tips.com/viewthread.cfm?qid=1804214
' https://wutils.com/wmi/root/microsoft/windows/storage/msft_physicaldisk/vbscript-samples.html
' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/stormgmt/msft-physicaldisk
Function getDiskType()
	'https://wutils.com/wmi/
	Dim oWMI, Instances, Instance
	
	'Get base WMI object, "." means computer name (local)
	Set oWMI = GetObject("WINMGMTS:\\.\ROOT\Microsoft\Windows\Storage")
	
	'Get instances of MSFT_PhysicalDisk - all instances of this class and derived classes 
	'Set Instances = oWMI.InstancesOf("MSFT_PhysicalDisk")
	
	'Get instances of MSFT_PhysicalDisk 
	Set Instances = oWMI.InstancesOf("MSFT_PhysicalDisk", 1)
	
	getDiskType = "Unspecified"
	'Enumerate instances  
	For Each Instance In Instances 
	  'Do something with the instance
	  Select Case Instance.MediaType
                Case 3
                    getDiskType = "HDD"
                Case 4
                    getDiskType = "SSD"
                Case 5
                    getDiskType = "SCM"
		Case 17
		    getDiskType = "NVMe SSD"
            End Select
	Next 'Instance
End Function

' Text to describe which type of hardware we'r on
Function getMaterielType()
	getMaterielType = "PC"
	if isTouchHardware() then
		getMaterielType = "Tablette"
	end if
End Function

' Test if this computer is a touch hardware
Function isTouchHardware()
	Dim res
	res = False
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PnPEntity")
	For Each objItem In colItems
    		If InStr(1, objItem.Description , "touch", 1) > 0 Then
			isTouchHardware = True
		End If
	Next
End Function

' Get CPU indice from cpu benchmark
Function getCPUindice()
	Dim res
	res = getCPUbenchmark(getCPUnameForCB())
	getCPUindice = res
End Function

' Get cpu indice from cpu bench mark html file
Function getCPUindiceFromHTML(myHTML, name)
	Set oRegExp2 = New RegExp
	oRegExp2.Pattern = ".*" & name & ".*<span class=.count.>([^<]+)</span>.*"
	Set matches2 = oRegExp2.Execute(myHTML)
	getCPUindiceFromHTML = matches2(0).SubMatches(0)
End Function

' Get cpu benchmark html file from cpu name
Function getCPUbenchmark(name)
	Set x = CreateObject("MSXML2.XMLHTTP")
	x.Open "Get", "https://www.cpubenchmark.net/cpu_lookup.php?cpu=" & urlEncode(name), False
	x.Send
	if x.Status = 200 then
'		Set fso = CreateObject("Scripting.FileSystemObject")
'		Set file = fso.OpenTextFile("res.txt", 2, True)
'		file.Write(x.responseText)
		getCPUbenchmark = getCPUindiceFromHTML(x.ResponseText, name)
	else
		MsgBox("cannot reach server, please ensure you'r connected to internet")
	end if
End Function

' Get special cpu format for injection in the cpu benchmark
Function getCPUnameForCB() 
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
	For Each objItem in colItems
	    getCPUnameForCB = objItem.Name
	    getCPUnameForCB = reReplaceAll(getCPUnameForCB, "\([^\)]+\)", "")
  	    getCPUnameForCB = reReplaceAll(getCPUnameForCB, "CPU ", "")
	    getCPUnameForCB = reReplaceAll(getCPUnameForCB, "@.*$", "")
	Next
End Function

' Display CPU useful informations to consumers
Function getCPU() 
	Dim res
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
	For Each objItem in colItems
	    res = res & objItem.Name & " L2 cache " & objItem.L2CacheSize & "Mo" & vbCrLf
	Next
	getCPU = res
End Function

' Cache memory on the system
Function getCacheMem()
	Dim res
	Set colItems = objWMIService.ExecQuery("Select * from Win32_CacheMemory",,48)
	For Each objItem in colItems    
	    res = res & objItem.Purpose & " " & objItem.InstalledSize & " Mo" & vbCrLf
	Next
	getCacheMem = res
End Function

' RAM memory installable on the system
Function getRAM()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray",,48)
	Dim res
	For Each objItem in colItems
		res = res & " maximum installable RAM " & objItem.MaxCapacity & " Ko in " & objItem.MemoryDevices & " slots " & vbCrLf
	Next
	getRam = res
End Function

' RAM memory installled on the system
Function getInstalledRAM()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
	Dim tot
	tot = 0
	For Each objItem in colItems
		tot = tot + objItem.Capacity
	Next
	tot = tot / 1000000
	Dim res
	res = res & "installed RAM quantity " & tot & " Mo"
	Set colItems2 = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
	Dim old
	For Each objItem in colItems2
		IF old=objItem.Speed THEN
		ELSE
			res = res & " " & objItem.Speed & "Mhz"
			old = objItem.Speed
		END IF
	Next
	getInstalledRAM = res
End Function

' RAM memory installled on the system
Function getInstalledRAMgo()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
	Dim tot
	tot = 0
	For Each objItem in colItems
		tot = tot + objItem.Capacity
	Next
	tot = tot / 1000000000
	getInstalledRAMgo = Round(tot, 3)
End Function

' Get softwares installed on this computer
Function getInstalledSoftware()
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Product",,48)
    Dim res
    For Each objItem in colItems
	res = res & objItem.Name & vbCrLf
    Next
    getInstalledSoftware = res
End Function

' Get connectivity infos
Function getConnectivity()
    Dim tot
    tot = 0
    Set colItems = objWMIService.ExecQuery("Select * from Win32_USBController",,48)
    For Each o in colItems
	tot = tot + 1
    Next
    getConnectivity = "" & tot & " USB ports" & vbCrLf
End Function

' Video card
Function getVideoCard()
    Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController",,48)
    Dim res
    For Each objItem in colItems
	res = res & objItem.Name & " " & objItem.AdapterRAM/1000000000 & " Go" & vbClRf
    Next
    getVideo = res
End Function

' Get disk space avaliable go
' WARNING : if there is any network volumes they will also be counted
Function getDiskSpaceGo()
    Dim tot
    tot = 0
    Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk",,48)
    For Each objItem in colItems
	If objItem.DriveType=3 THEN
		tot = tot + (objItem.Size/1000000000)
	END IF
    Next 
    getDiskSpaceGo = tot
END FUNCTION

' Get disk information
Function getDiskInfos() 
    Dim res
    Set colItems = objWMIService.ExecQuery("Select * from Win32_IDEController",,48)
    Dim dCount
    dCount = 0
    For Each objItem in colItems
	dCount = dCount + 1
    Next
    res = res & "Disk slots " & dCount & vbClRf
    res = res & ", amount of space : " & Round(getDiskSpaceGo(), 2) & "Go" & vbClRf
    getDiskInfos = res
End Function


' url encode some text
' space become +
Function urlEncode(text)
	urlEncode = reReplaceAll(text, " ", "+")
End Function

' Get status of the bluetooth Pres, Abs, KO
Function bluetoothSupported()
	bluetoothSupported = "Abs"
	Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkProtocol")
	For Each objItem in colItems
	    If InStr(objItem.Name, "Bluetooth") Then
		bluetoothSupported = "Pres"
		Exit For
	    End If
	Next
End Function

' Get screen resolution
Function getScreenResolutionPx()
	Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_VideoController" )
	For Each objItem In colItems
		getScreenResolutionPx = objItem.CurrentHorizontalResolution & " x " & objItem.CurrentVerticalResolution
	Next
End Function

' get current date eg 14/02/2021 10:00
Function curDate()
	Dim dt
	dt=now
	curDate = day(dt) & "/" & month(dt) & "/" & year(dt) & " " & hour(dt) & ":" & minute(dt)
End Function

' Format : MANUFACT + MODELE + CPU FREQ + RAM
Function getNomComplet()
	Dim man, model
	Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)
	For Each objItem in colItems
		man = objItem.Vendor
		model = objItem.Version
	Next
	getNomComplet = man & " " & model & " stockage " & Round(getDiskSpaceGo()) & "Go RAM " & Round(getInstalledRAMgo()) & "Go"
End Function

' Redim preserve on multidim arrays without out of range exception
' array to copy
' w new width
' h new height
Function ReDimPreserve(arr, ph, pw)
    Dim h, w
    h = Max(ph, UBound(arr, 1))
    w = Max(pw, UBound(arr, 2))
    ReDim newArr(h, w)
    y = 0
    Do While y<UBound(arr, 1)
        x = 0
        Do While x<UBound(arr, 2)
            If x<w and y<h then
                    newArr(y, x) = arr(y, x)
                else
                    newArr(y, x) = 0
                end if
            x=x+1
        Loop
        y = y + 1
    Loop
    ReDim arr(h, w)

    y = 0
    Do While y<UBound(arr, 1)
        x = 0
        Do While x<UBound(arr, 2)
            arr(y, x) = newArr(y, x)
            x=x+1
        Loop
        y = y + 1
    Loop
End Function

Function Min(x, y)
    If x < y Then Min = x Else Min = y
End Function

Function Max(x, y)
    If x > y Then Max = x Else Max = y
End Function

' Returns True if string in parameter is upper case
Function IsUpper(s)
    With CreateObject("VBScript.RegExp")
        .Pattern = "^[^a-z]*$"
        IsUpper = .test(s)
    End With
End Function

' regexp replace
Function reReplace(strString, strPattern, strReplace)
    Dim oRegExp
    Set oRegExp = New RegExp
    oRegExp.Pattern = strPattern
    reReplace = oRegExp.Replace(strString, strReplace)
End Function

' regexp replace all
Function reReplaceAll(strString, strPattern, strReplace)
    Dim oRegExp
    Set oRegExp = New RegExp
    oRegExp.Global = True
    oRegExp.Pattern = strPattern
    reReplaceAll = oRegExp.Replace(strString, strReplace)
End Function

' take a char in a string a returns its value
Function charToNumberValue(s)
    IF IsUpper(s) THEN
        charToNumberValue = Asc(s)  - 65
    ELSE
        charToNumberValue = Asc(s) - 97
    END IF
End Function

sub assert( boolExpr, str )
    if not boolExpr then
        Err.Raise vbObjectError + 99999, , str
    end if
end sub

		
		'*******************************************************************************
'*     VBScript Binary Functions 
'*     http://chris.wastedhalo.com
'*
'*     VBScript doesn't have the greatest support for Binary Operations
'*     so I've created these functions for myself and decided to share them
'*     with anyone who is interested.  One of my issues with VBScript is that
'*     it's constantly changing the sub-type of my variables when doing bit-wise
'*     operations.  It also throws up overflow errors when you try and mess the
'*     sign bit.  I have figured out a few tricks to prevent these problems and
'*     have incorporated them into these functions.
'*
'*     These functions all detect the sub-type of your variables and will preserve
'*     them.  They will work with Long Integers, Integers and Bytes but you
'*     Must set the sub-type of your variables using CLng(), CInt() or CByte()
'*     before hand to get the correct results.
'*
'*     Long Integer = 32 Bits - Set Sub-Type with CLng()
'*          Integer = 16 Bits - Set Sub-Type with CInt()
'*             Byte = 8 Bits  - Set Sub-Type with CByte()
'*
'*     I tried to keep these as simple as possible while preserving sub-types,
'*     working around VBScripts quirks and still having some error checking.
'*     If you have any comments or suggestions please post them on my blog.
'*
'*     Do what you like with these functions.  All I ask is that you keep a 
'*     link back to my site included with them if they're shared or reposted.
'*                         chris.wastedhalo.com
'******************************************************************************* 
 
'******************************************************************************* 
'*     GetBit(AnyNumber, BitNumberToCheck)
'*         Returns True if bit is a 1, False if bit is a 0
'*         Sub-Type does not matter
'*******************************************************************************
Function GetBit(pValue, pBit)
     Dim BitMask
     If pBit > 32 Then Err.Raise 6 ' Overflow (Bit number too high)
     If pBit < 32 Then BitMask = 2 ^ (pBit - 1) Else BitMask = "&H80000000"
     GetBit = CBool(pValue AND BitMask)
End Function
 
'******************************************************************************* 
'*      SetBit(AnyNumber, BitNumberToChange, ChangeBitTo)
'*          Returns a new number with your bit changed.
'*          For the pNewValue argument you can use True/False or (1 or 0) 
'******************************************************************************* 
Function SetBit(pValue, pBit, pNewValue)
    Dim NewValue, BitMask, vType
    If pBit > 32 Then Err.Raise 6 ' Bit number too high
    If pBit < 32 Then BitMask = 2 ^ (pBit - 1) Else BitMask = "&H80000000"
    vType = VarType(pValue)
    If vType <> vbLong And vType <> vbInteger And vType <> vbByte Then Err.Raise 13
    If pNewValue Then
        NewValue = CLng(pValue Or BitMask) 
    Else
        NewValue = CLng(pValue And Not BitMask)
    End If
    Select Case vType
        Case vbLong: SetBit = CLng(NewValue)
        Case vbInteger: SetBit = CInt("&H"+ Hex(NewValue And "&HFFFF")) 
        Case vbByte: SetBit = cByte(NewValue And "&HFF")
    End Select
End Function
 
'******************************************************************************* 
'*      ToggleBit(AnyNumber, BitNumberToToggle)
'*          Returns a new number with your bit toggled.
'******************************************************************************* 
Function ToggleBit(pValue, pBit)
    Dim BitMask
    If pBit > 32 Then Err.Raise 6 ' Bit number too high
    If pBit < 32 Then BitMask = 2 ^ (pBit - 1) Else BitMask = "&H80000000"
    Select Case VarType(pValue)
        Case vbLong: ToggleBit = pValue XOR BitMask
        Case vbInteger: ToggleBit = CInt("&H"+ Hex((pValue XOR BitMask) And "&HFFFF"))
        Case vbByte: ToggleBit = CByte((pValue XOR BitMask) And "&HFF")
        Case Else: Err.Raise 13 ' Not a supported type 
    End Select 
End Function
 
'*******************************************************************************
'*     ExtractBits(AnyNumber, BitStartPosition, NumberOfBits)
'*         Returns the decimal value of the extracted bits.
'*******************************************************************************
Function ExtractBits(pValue, pStartPos, pBits)
    Dim BitMask, tmpMask, i, NewValue
    For i = pStartPos - pBits + 1 To pStartPos
        If i < 32 Then tmpMask = 2 ^ (i - 1) Else tmpMask = "&H80000000"
        BitMask = BitMask Or tmpMask
    Next
    NewValue = CLng(pValue And BitMask)
    If NewValue And "&H80000000" Then tmpMask = pBits Else tmpMask = 0
    NewValue = (NewValue And "&H7FFFFFFF") / 2 ^ (pStartPos - pBits)
    If tmpMask Then
        If tmpMask < 32 Then BitMask = 2 ^ (tmpMask - 1) Else BitMask = "&H80000000"
        NewValue = NewValue Or BitMask
    End If
    ExtractBits = CLng(NewValue)
End Function
 
'*******************************************************************************
'*     LeftShift(AnyNumber, BitsToShiftBy)
'*         Returns a new number with bits shifted. 0's are shifted in from the
'*         right, bits will fall off on the left.
'*******************************************************************************
Function LeftShift(pValue, pShift)
    Dim NewValue, PrevValue, i
    PrevValue = pValue
    For i = 1 to pShift
        Select Case VarType(pValue)
            Case vbLong
                NewValue = (PrevValue And "&H3FFFFFFF") * 2
                If PrevValue And "&H40000000" Then NewValue = NewValue Or "&H80000000"
                NewValue = CLng(NewValue)
            Case vbInteger
                NewValue = (PrevValue And "&H3FFF") * 2
                If PrevValue And "&H4000" Then NewValue = NewValue Or "&H8000"
                NewValue = CInt("&H"+ Hex(NewValue))
            Case vbByte
                NewValue = CByte((PrevValue And "&H7F") * 2)
            Case Else: Err.Raise 13 ' Not a supported type 
        End Select
        PrevValue = NewValue
    Next
    LeftShift = NewVAlue
End Function
 
'*******************************************************************************
'*     RollLeft(AnyNumber, BitsToShiftBy)
'*         Returns a new number with bits shifted
'*         Bits are shifted to the left. Bits that that fall off 
'*         get rolled over to the right.
'*******************************************************************************
Function RollLeft(pValue, pRoll)
    Dim NewValue, PrevValue, i
    PrevValue = pValue
    For i = 1 to pRoll
        Select Case VarType(pValue)
            Case vbLong
                NewValue = (PrevValue And "&H3FFFFFFF") * 2
                If PrevValue And "&H40000000" Then NewValue = NewValue Or "&H80000000"
                If PrevValue And "&H80000000" Then NewValue = NewValue Or "&H1"
                NewValue = CLng(NewValue)
            Case vbInteger
                NewValue = (PrevValue And "&H3FFF") * 2
                If PrevValue And "&H4000" Then NewValue = NewValue Or "&H8000"
                If PrevValue And "&H8000" Then NewValue = NewValue Or "&H1"
                NewValue = CInt("&H"+ Hex(NewValue))
            Case vbByte
                NewValue = (PrevValue And "&H7F") * 2
                If PrevValue And "&H80" Then NewValue = NewValue Or "&H1"
                NewValue = CByte(NewValue)
            Case Else: Err.Raise 13 ' Not a supported type 
        End Select
        PrevValue = NewValue
    Next
    RollLeft = NewVAlue
End Function
 
'*******************************************************************************
'*     RightShift(AnyNumber, BitsToShiftBy)
'*         Returns a new number with bits shifted
'*         0's are shifted in from the left. Bits will fall off on the right.
'*******************************************************************************
Function RightShift(pValue, pShift)
    Dim NewValue, PrevValue, i
    PrevValue = pValue
    For i = 1 to pShift
        Select Case VarType(pValue)
            Case vbLong
                NewValue = Int((PrevValue And "&H7FFFFFFF") / 2)
                If PrevValue And "&H80000000" Then NewValue = NewValue Or "&H40000000"
                NewValue = CLng(NewValue)
            Case vbInteger
                NewValue = Int((PrevValue And "&H7FFF") / 2)
                If PrevValue And "&H8000" Then NewValue = NewValue Or "&H4000"
                NewValue = CInt(NewValue)
            Case vbByte
                NewValue = CByte(PrevValue / 2)
            Case Else: Err.Raise 13 ' Not a supported type
        End Select
        PrevValue = NewValue
    Next
    RightShift = PrevValue
End Function
 
'*******************************************************************************
'*     SignedRightShift(AnyNumber, BitsToShiftBy)
'*         Returns a new number with bits shifted
'*         The sign bit is copied and shifted in from the left. 
'*         Bits will fall off on the right.
'*******************************************************************************
Function SignedRightShift(pValue, pShift)
    Dim NewValue, PrevValue, i
    PrevValue = pValue
    For i = 1 to pShift
        Select Case VarType(pValue)
            Case vbLong
                NewValue = Int((PrevValue And "&H7FFFFFFF") / 2)
                If PrevValue And "&H80000000" Then NewValue = NewValue Or "&HC0000000"
                NewValue = CLng(NewValue)
            Case vbInteger
                NewValue = Int((PrevValue And "&H7FFF") / 2)
                If PrevValue And "&H8000" Then NewValue = NewValue Or "&HC000"
                NewValue = CInt("&H"+ Hex(NewValue))
            Case vbByte
                NewValue = Int(PrevValue / 2)
                If PrevValue And "&H80" Then NewValue = NewValue Or "&HC0"
                NewValue = CByte(NewValue)
            Case Else: Err.Raise 13 ' Not a supported type
        End Select
        PrevValue = NewValue
    Next
    SignedRightShift = PrevValue
End Function
 
'*******************************************************************************
'*     RollRight(AnyNumber, BitsToShiftBy)
'*         Returns a new number with bits shifted
'*         Bits are shifted to the right. Bits that fall off 
'*         get rolled over to the left.
'*******************************************************************************
Function RollRight(pValue, pRoll)
    Dim NewValue, PrevValue, i
    PrevValue = pValue
    For i = 1 to pRoll
        Select Case VarType(pValue)
            Case vbLong
                NewValue = Int((PrevValue And "&H7FFFFFFF") / 2)
                If PrevValue And "&H80000000" Then NewValue = NewValue Or "&H40000000"
                If PrevValue And "&H1" Then NewValue = NewValue Or "&H80000000"
                NewValue = CLng(NewValue)
            Case vbInteger
                NewValue = Int((PrevValue And "&H7FFF") / 2)
                If PrevValue And "&H8000" Then NewValue = NewValue Or "&H4000"
                If PrevValue And "&H1" Then NewValue = NewValue Or "&H8000"
                NewValue = CInt("&H"+ Hex(NewValue))
            Case vbByte
                NewValue = Int(PrevValue / 2)
                If PrevValue And "&H1" Then NewValue = NewValue Or "&H80"
                NewValue = CByte(NewValue)
            Case Else: Err.Raise 13 ' Not a supported type
        End Select
        PrevValue = NewValue
    Next
    RollRight = PrevValue
End Function
 
'*******************************************************************************
'*     bMask(BitNumber)
'*         Returns a number with all bits set to 0 except for the specified bit
'*******************************************************************************
Function bMask(pBit)
    If pBit < 32 Then bMask = 2 ^ (pBit - 1) Else bMask = "&H80000000"
End Function
 
'*******************************************************************************
'*     Dec2Bin(AnyNumber)
'*         Returns a string representing the number in binary.
'*******************************************************************************
Function Dec2Bin(pValue)
    Dim TotalBits, i
    Select Case VarType(pValue)
        Case vbLong: TotalBits = 32
        Case vbInteger: TotalBits = 16
        Case vbByte: TotalBits = 8
        Case Else: Err.Raise 13 ' Not a supported type
    End Select
    For i = TotalBits to 1 Step -1
        If pValue And bMask(i) Then Dec2Bin = Dec2Bin + "1" Else Dec2Bin = Dec2Bin + "0"
    Next
End Function
 
'*******************************************************************************
'*     Bin2Dec(BinaryString)
'*         Returns the decimal value of a string of binary.
'*******************************************************************************
Function Bin2Dec(pBinString)
    Dim Binary, i
    Binary = Trim(Right(pBinString, Len(pBinString) - Instr(pBinString,"1") + 1))
    If Len(Binary) > 32 Then Err.Raise 6' Overflow
    For i = 1 To Len(Binary)
        Select Case Mid(Binary, i, 1)
            Case "1": Bin2Dec = Bin2Dec Or bMask(Len(Binary) - i + 1)
            Case "0": 'Do Nothing
            Case Else: Err.Raise 13 ' Not 1 or 0 (Type Mismatch Error) 
        End Select 
    Next
End Function
