'
' This is exposed to consumer as a one click reports generation tool
'
MsgBox("L'inspecteur va chercher après que vous ayez validé")

' load a library when network is avaliable
' it takes filename in the remote location and load the lib
' or abort with error
Function load(filename)

    ' create a cache folder in the local path for execution offline
    Const folder = ".cache"
    Const Overwrite = True
    Dim oFSO
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    If Not oFSO.FolderExists(folder) Then
        Set f = oFSO.CreateFolder(folder)
        Dim rFSO
        Set rFSO = CreateObject("Scripting.FileSystemObject")
	Set fdir = oFSO.GetFolder(folder)
	fdir.Attributes = 2   
    End If

    ' fetch library from remote if network is avaliable
    Dim lib
    Dim o
    Dim fname
    fname = folder & "\" & filename
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set o = CreateObject("MSXML2.XMLHTTP")
    o.open "GET", "https://raw.githubusercontent.com/emmausConnect/inspector/main/" & filename, False
    o.send
    IF o.Status = 200 THEN
    	lib = o.responseText
	Set file = fso.OpenTextFile(fname, 2, True)
	file.Write(lib)
    ELSE
	IF fso.FileExists(fname) THEN
		Set file = fso.OpenTextFile(fname, 1)
		lib = file.ReadAll
	ELSE
		MsgBox "Cannot load library please run inspector with internet connection (at least the first time)" & vbClRf & "Aborting ...", 16
		wscript.Quit
	END IF
    END IF
    executeGlobal lib
End Function

load("lib.vbs")
load("libExcelReports.vbs")


Dim f
f = "reports.xlsx"
Set o = sheetOpenOrCreate(f)
	
sheetUpdateOrNewEntryFromThisPC(o("sheet"))
sheetAutoFit(o("sheet"))
sheetWrite o, f
o("objExcel").Quit

MsgBox("L'inspection est finie")
