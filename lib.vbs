'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Library to extract informations from computer
'It is not oriented to diagnose computers but give an overview of it's features
'doc: https://www.activexperts.com/admin/scripts/wmi/
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strComputer
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
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
