'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Library to extract informations from computer
'It is not oriented to diagnose computers but give an overview of it's features
'doc: https://www.activexperts.com/admin/scripts/wmi/
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strComputer
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

' Get capacity remaining (from it's design capacity)
' Windows 2000 and Windows 98 doivent avoir activé APM
' https://docs.microsoft.com/fr-fr/windows/win32/cimwin32prov/win32-battery
Function getBatteryCapResid() 

	' [Windows Vista;[
	Dim objItem
	getBatteryCapResid = "N.A."
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Battery",,48)
	For Each objItem in colItems
		if Not objItem.FullChargeCapacity=0 then
			getBatteryCapResid = objItem.FullChargeCapacity
		end if
	Next

End Function

' Get a string describing the mouse currently connected to computer.
Function getMouseStr()

	getMouseStr = ""

	' [Windows Vista;[
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PointingDevice",,48)
	For Each objItem in colItems
		if objItem.PointingType=3 or objItem.PointingType=9 then
			if objItem.ConfigManagerErrorCode=0 then
				getMouseStr = "Pres"
				exit function
			else
				getMouseStr = "KO"
			end if
		end if
	Next

End function

' Test if a physical keyboard is present.
' Peripherical keyboard are not taken in account.
Function getKeyboard()

	getKeyboard = "Abs"

	' [Windows Vista;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Keyboard")
	For Each objItem in colItems
		IF InStr(objItem.Description,"USB") THEN
		ELSE
			IF objItem.ConfigManagerErrorCode=0 THEN
				getKeyboard = "Pres"
				exit function
			ELSE
				getKeyboard = "KO"
			end IF
		end if
	Next

End Function

' Try to guess a string to describe alim chargeur.
Function guessAlimChargeur()
	guessAlimChargeur = getBatteryCapResid()
	if not(guessAlimChargeur = "N.A.") then
		guessAlimChargeur = ""
	end if		
End Function

' Get amount of time that this computer can live on (in min)
' Windows 2000 and Windows 98 doivent avoir activé APM
' https://docs.microsoft.com/fr-fr/windows/win32/cimwin32prov/win32-battery
Function getBatteryAmountTimeExpected()

	getBatteryAmountTimeExpected = "N.A."
	
	' [Windows Vista;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Battery",,48)
	For Each objItem in colItems
		getBatteryAmountTimeExpected = objItem.ExpectedLife
	Next
End Function

' Test presence of CDROM on this computer
Function getCDROMinfos()

	getCDROMinfos = "Abs"

	' [Windows Vista;[
	Dim item
	Set items = objWMIService.ExecQuery("Select * From Win32_CDROMDrive",,48)
	for each item in items
		if item.ConfigManagerErrorCode=0 then
			getCDROMinfos = "Pres"
			exit function
		else
			getCDROMinfos = "KO"
		end if
	next

End Function

' Get status of the bluetooth Pres, Abs, KO
' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/hh968170(v=vs.85)
Function bluetoothSupported()
	
	bluetoothSupported = "Abs"

	' 0. [Windows Vista;[
	Dim objItem, item
	Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter")
	For Each objItem in colItems
		If objItem.AdapterTypeID = 9 And objItem.PhysicalAdapter then ' wireless card
			Set oRegExp2 = New RegExp
			oRegExp2.Pattern = ".*[Bb]lue[Tt]ooth.*"
			if oRegExp2.Test(objItem.ProductName) Or oRegExp2.Test(objItem.Description) then
				if objItem.ConfigManagerUserConfig=0 then
					bluetoothSupported = "Pres"
					exit function
				else
					bluetoothSupported = "KO"
				end if
			end if
		end if
	Next

	' 1. [Windows 8;[
	Err.Clear
	On Error Resume Next
	Set newSpace = GetObject("winmgmts:\\" & strComputer & "\root\StandardCimv2")
	If Err.Number=0 Then
		Set items = newSpace.ExecQuery("select Name, InterfaceName, InterfaceType, NdisPhysicalMedium from MSFT_NetAdapter where ConnectorPresent=1")
		for each item in items
			if item.NdisPhysicalMedium=10 then
				if item.ConfigManagerUserConfig=0 then
					bluetoothSupported = "Pres"
					exit function
				else
					bluetoothSupported = "KO"
				end if
			end if
		next
	end if
	On Error Goto 0

End Function

' Check if ethernet port is present
' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/hh968170(v=vs.85)
' https://askcodez.com/determiner-le-type-de-carte-reseau-via-wmi.html
Function ethernetPort() 

	ethernetPort="Abs"

	' 0. [Windows Vista;[
	Dim objItem, item
	Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter")
	For Each objItem in colItems
		If objItem.AdapterTypeID=0 And objItem.PhysicalAdapter then ' Ethernet 802,3 device
			ethernetPort="Pres"			
		end if
	Next

	' 1. [Windows 8;[
	On Error Resume Next
	Err.Clear
	Set newSpace = GetObject("winmgmts:\\" & strComputer & "\root\StandardCimv2")
	If Err.Number=0 Then
		Set items = newSpace.ExecQuery("select Name, InterfaceName, InterfaceType, LinkTechnology, NdisPhysicalMedium from MSFT_NetAdapter where ConnectorPresent=1")
		For Each item in items
			if item.InterfaceType=6 and ( item.NdisPhysicalMedium=0 or item.NdisPhysicalMedium=14 ) then
				ethernetPort="Pres"
			end if
		Next
	end if
	On Error Goto 0

End Function

' Get a string describing the type of disk used inside this computer
' https://www.tek-tips.com/viewthread.cfm?qid=1804214
' https://wutils.com/wmi/root/microsoft/windows/storage/msft_physicaldisk/vbscript-samples.html
' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/stormgmt/msft-physicaldisk
Function getDiskType()
	
	' [Windows 8;[ 
	' Get information from physical disk
	'https://wutils.com/wmi/
	Dim oWMI, Instances, Instance
	'Get base WMI object, "." means computer name (local)
	Set oWMI = GetObject("WINMGMTS:\\.\ROOT\Microsoft\Windows\Storage")
	On Error Resume Next
	Err.Clear
	'Get instances of MSFT_PhysicalDisk 
	Set Instances = oWMI.InstancesOf("MSFT_PhysicalDisk", 1)
	if Err.Number=0 then
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
	end if
	
End Function

' Text to describe which type of hardware we'r on
Function getMaterielType()
	getMaterielType = "PC"
	if not(guessAlimChargeur = "N.A.") then
		getMaterielType = "PC Portable"
	end if
	if isTouchHardware() then
		getMaterielType = "Tablette"
	end if
End Function

' Test if this computer is a touch hardware
Function isTouchHardware()
	isTouchHardware = False
	
	' [Windows Vista;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PnPEntity")
	For Each objItem In colItems
    		If InStr(1, objItem.Description , "touch", 1) > 0 Then
			isTouchHardware = True
		End If
	Next
	
	' [Windows Vista;[
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PointingDevice",,48)
	For Each objItem in colItems
		if objItem.PointingType=7 or objItem.PointingType=8 then
			isTouchHardware = True
		end if
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
	
	' [Windows Vista;[
	Dim objItem
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
	
	' [Windows Vista;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
	For Each objItem in colItems
	    getCPU = getCPU & objItem.Name & " L2 cache " & objItem.L2CacheSize & "Mo" & vbCrLf
	Next
	
End Function

' Cache memory on the system
Function getCacheMem()
	
	' [Windows Vista;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_CacheMemory",,48)
	For Each objItem in colItems    
	    getCacheMem = getCacheMem & objItem.Purpose & " " & objItem.InstalledSize & " Mo" & vbCrLf
	Next

End Function

' RAM memory installable on the system
Function getRAM()
	
	' [Windows Vista;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray",,48)
	For Each objItem in colItems
		getRAM = getRAM & " maximum installable RAM " & objItem.MaxCapacity & " Ko in " & objItem.MemoryDevices & " slots " & vbCrLf
	Next

End Function

' RAM memory installled on the system
Function getInstalledRAM()
	
	' [Windows Vista;[
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
	Dim tot
	tot = 0
	For Each objItem in colItems
		tot = tot + objItem.Capacity
	Next
	tot = tot / 1000000
	getInstalledRAM = getInstalledRAM & "installed RAM quantity " & tot & " Mo"
	Set colItems2 = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
	Dim old, objItem
	For Each objItem in colItems2
		IF old=objItem.Speed THEN
		ELSE
			getInstalledRAM = getInstalledRAM & " " & objItem.Speed & "Mhz"
			old = objItem.Speed
		END IF
	Next
	
End Function

' RAM memory installled on the system
Function getInstalledRAMgo()
	
	' [Windows Vista;[
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
	Dim tot, objItem
	tot = 0
	For Each objItem in colItems
		tot = tot + objItem.Capacity
	Next
	tot = tot / 1000000000
	getInstalledRAMgo = Round(tot, 3)
	
End Function

' Get softwares installed on this computer
Function getInstalledSoftware()

	' [Windows XP;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Product",,48)
	For Each objItem in colItems
		getInstalledSoftware = getInstalledSoftware & objItem.Name & vbCrLf
	Next
	
End Function

' Get the number of usb ports on the system
Function getNumberUSBports() 
	
	' [Windows Vista;[
	Dim o
    	getNumberUSBports = 0
    	Set colItems = objWMIService.ExecQuery("Select * from Win32_USBController",,48)
    	For Each o in colItems
		getNumberUSBports = getNumberUSBports + 1
    	Next	
	
End Function

' Get connectivity infos
Function getConnectivity()
    getConnectivity = "" & getNumberUSBports() & " USB ports" & vbCrLf
End Function

' Video card
Function getVideoCard()
	
	' [Windows Vista;[
    	Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController",,48)
	Dim res, objItem
    	For Each objItem in colItems
		res = res & objItem.Name & " " & objItem.AdapterRAM/1000000000 & " Go" & vbClRf
    	Next
    	getVideo = res
	
End Function

' Get disk space avaliable go
' WARNING : if there is any network volumes they will also be counted
Function getDiskSpaceGo()
	
	' [Windows Vista;[
	Dim tot, objItem
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
	
	' [Windows Vista;[
    	Dim res
    	Set colItems = objWMIService.ExecQuery("Select * from Win32_IDEController",,48)
	Dim dCount, objItem
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

' Get screen resolution in pixels (the amount of pixel in the current screen).
' https://docs.microsoft.com/fr-fr/windows/win32/cimwin32prov/win32-videocontroller
' https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/cim-desktopmonitor
Function getScreenResolutionPx()
	Dim w(1), h(1)

	' 0. [Windows Vista;[ Get current dimensions in pixels
	Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_VideoController" )
	For Each objItem In colItems
		w(0) = objItem.CurrentHorizontalResolution
		h(0) = objItem.CurrentVerticalResolution
	Next

	' 1. [Windows Vista;[ Get dimensions in pixels from supported source mode
	Set specialWMIService = GetObject("winmgmts:\\.\root\WMI")
	Set colItems = specialWMIService.ExecQuery("Select * From WmiMonitorListedSupportedSourceModes")
	For Each objItem in colItems
		w(1) = objItem.MonitorSourceModes(0).HorizontalActivePixels 
		h(1) = objItem.MonitorSourceModes(0).VerticalActivePixels
	next

	Dim wmax, hmax
	wmax = -1
	hmax = -1
	for a = 0 to UBound(w)
		if wmax=-1 Or hmax=-1 Or ( hmax < h(a) And wmax < w(a)) then
			wmax = w(a)
			hmax = h(a)
			getScreenResolutionPx = wmax & "x" & hmax
		end if
	next
End Function

' get current date eg 14/02/2021
Function curDate2()
	Dim dt
	dt=now
	curDate2 = day(dt) & "/" & month(dt) & "/" & year(dt)
End Function

' get current date eg 14/02/2021 10:00
Function curDate()
	Dim dt
	dt=now
	curDate = day(dt) & "/" & month(dt) & "/" & year(dt) & " " & hour(dt) & ":" & minute(dt)
End Function

' computer model : MANUFACT
Function getMarque()
	
	' [Windows Vista;[
	Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)
	For Each objItem in colItems
		getMarque = objItem.Vendor
	Next
	
End Function

' computer model : MANUFACT + MODELE
Function getModel()
	
	' [Windows Vista;[
	Dim objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)
	For Each objItem in colItems
		getModel = objItem.Version
	Next
	
End Function

' Format : MANUFACT + MODELE + CPU FREQ + RAM
Function getNomComplet()
	
	' [Windows Vista;[
	Dim man, model, objItem
	Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)
	For Each objItem in colItems
		man = objItem.Vendor
		model = objItem.Version
	Next
	getNomComplet = getMarque() & " " & getModel() & " stockage " & Round(getDiskSpaceGo()) & "Go RAM " & Round(getInstalledRAMgo()) & "Go"
	
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

Dim ConfigManagerErrorCodeInfo
Set ConfigManagerErrorCodeInfo = CreateObject("Scripting.Dictionary")
ConfigManagerErrorCodeInfo.Add 0, "This device is working properly. "
ConfigManagerErrorCodeInfo.Add 1, "This device is not configured correctly. "
ConfigManagerErrorCodeInfo.Add 2, "Windows cannot load the driver for this device. "
ConfigManagerErrorCodeInfo.Add 3, "The driver for this device might be corrupted, or your system may be running low on memory or other resources. "
ConfigManagerErrorCodeInfo.Add 4, "This device is not working properly. One of its drivers or your registry might be corrupted. "
ConfigManagerErrorCodeInfo.Add 5, "The driver for this device needs a resource that Windows cannot manage. "
ConfigManagerErrorCodeInfo.Add 6, "The boot configuration for this device conflicts with other devices. "
ConfigManagerErrorCodeInfo.Add 7, "Cannot filter. "
ConfigManagerErrorCodeInfo.Add 8, "The driver loader for the device is missing. "
ConfigManagerErrorCodeInfo.Add 9, "This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly. "
ConfigManagerErrorCodeInfo.Add 10, "This device cannot start. "
ConfigManagerErrorCodeInfo.Add 11, "This device failed. "
ConfigManagerErrorCodeInfo.Add 12, "This device cannot find enough free resources that it can use. "
ConfigManagerErrorCodeInfo.Add 13, "Windows cannot verify this device's resources. "
ConfigManagerErrorCodeInfo.Add 14, "This device cannot work properly until you restart your computer. "
ConfigManagerErrorCodeInfo.Add 15, "This device is not working properly because there is probably a re-enumeration problem. "
ConfigManagerErrorCodeInfo.Add 16, "Windows cannot identify all the resources this device uses. "
ConfigManagerErrorCodeInfo.Add 17, "This device is asking for an unknown resource type. "
ConfigManagerErrorCodeInfo.Add 18, "Reinstall the drivers for this device. "
ConfigManagerErrorCodeInfo.Add 19, "Failure using the VxD loader. "
ConfigManagerErrorCodeInfo.Add 20, "Your registry might be corrupted. "
ConfigManagerErrorCodeInfo.Add 21, "System failure: Try changing the driver for this device. If that does not work, see your hardware documentation. Windows is removing this device. "
ConfigManagerErrorCodeInfo.Add 22, "This device is disabled. "
ConfigManagerErrorCodeInfo.Add 23, "System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation. "
ConfigManagerErrorCodeInfo.Add 24, "This device is not present, is not working properly, or does not have all its drivers installed. "
ConfigManagerErrorCodeInfo.Add 25, "Windows is still setting up this device. "
ConfigManagerErrorCodeInfo.Add 26, "Windows is still setting up this device. "
ConfigManagerErrorCodeInfo.Add 27, "This device does not have valid log configuration. "
ConfigManagerErrorCodeInfo.Add 28, "The drivers for this device are not installed. "
ConfigManagerErrorCodeInfo.Add 29, "This device is disabled because the firmware of the device did not give it the required resources. "
ConfigManagerErrorCodeInfo.Add 30, "This device is using an Interrupt Request (IRQ) resource that another device is using. "
ConfigManagerErrorCodeInfo.Add 31, "This device is not working properly because Windows cannot load the drivers required for this device. "


