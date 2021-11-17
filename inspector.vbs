

Const F_CACHE = ".cache"

' get name of a file in the cache
Function cache(name)
  Set oFSO = CreateObject("Scripting.FileSystemObject")

  If Not oFSO.FolderExists(F_CACHE) Then
    Set f = oFSO.CreateFolder(F_CACHE)
    Dim rFSO
    Set rFSO = CreateObject("Scripting.FileSystemObject")
    Set fdir = oFSO.GetFolder(F_CACHE)
    fdir.Attributes = 2   
  End If

  cache = F_CACHE & "\" & name

End Function

' Fetch libs online first if needed
Function fetchAllFirstIfNeeded()
  Set oFSO = CreateObject("Scripting.FileSystemObject")

  If Not oFSO.FolderExists(F_CACHE) Then
    Dim arr(4), x
    arr(0) = "lib.vbs"
    arr(1) = "libReports.vbs"
    arr(2) = "libExcelReports.vbs"
    arr(3) = "libOpenOfficeReports.vbs"
    arr(4) = "libCsvReports.vbs"
    x = 0
    Do While x<=UBound(arr)
      IF VarType(fetch(arr(x)))=0 THEN
        MsgBox "Cannot load library please run inspector without internet connection (at least the first time)" & vbCrLf & "Aborting ...", 16
        wscript.Quit
      END IF
      x = x + 1
    Loop
  End IF
End Function

' Fetch resource online and put in cache if possible
' Return the text of the resource or nothing in case of error
Function fetch(filename)
  Dim fname
  fname = cache(filename)
  Set o = CreateObject("MSXML2.XMLHTTP")
  o.open "GET", "https://raw.githubusercontent.com/emmausConnect/inspector/main/" & filename, False
  o.setRequestHeader "Accept", "application/vnd.github.v3.raw" 
  
  o.send
  IF o.Status = 200 THEN
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(fname, 2, True)
    file.Write(o.responseText)
    fetch = o.responseText
  END IF
End Function

Dim fetchProba
Randomize
fetchProba = Rnd()
' function fetch rarely eg 1 on 200 use of the function
Function fetchRarely(filename)
    if fetchProba < 0.005 THEN
      fetchRarely = fetch(filename)
    end if
End Function

' load a library when network is avaliable
' it takes filename in the remote location and load the lib
' or abort with error
Function load(filename)
  ' fetch library from remote if network is avaliable
  Dim lib, fname
  lib = fetchRarely(filename)
  IF VarType(lib)=0 THEN
    fname = cache(filename)
    Set fso = CreateObject("Scripting.FileSystemObject")
    IF fso.FileExists(fname) THEN
      Set file = fso.OpenTextFile(fname, 1)
      lib = file.ReadAll
    ELSE
      MsgBox "Libraries are not loaded please retry with internet connection please", 16
      wscript.Quit
    END IF
  END IF
  executeGlobal lib
End Function

Dim outputFilename, outputFile
outputFilename = "reports"

fetchAllFirstIfNeeded()

load("lib.vbs")
load("libReports.vbs")
loadPreferedBackend()

outputFile = getOutputFile(outputFilename)



MsgBox("L'inspecteur va chercher après que vous validez")

Set o = sheetOpenOrCreate(outputFile)

If TypeName(o("sheet")) = "Worksheet" THEN
  Set o("sheet") = sheetUpdateOrNewEntryFromThisPC(o("sheet"))
ELSE
  o("sheet") = sheetUpdateOrNewEntryFromThisPC(o("sheet"))
END IF
sheetAutoFit(o("sheet"))
sheetWrite o, outputFile
sheetClose(o)

MsgBox("Inspection terminée")










