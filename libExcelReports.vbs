' Functions to interact with excel document (old format)
Dim positions 
Set positions = getPositionsIndex()

' Get order of items in the row
Function getPositionsIndex()
    Dim positions
    Set positions = CreateObject("Scripting.Dictionary")
    positions.Add "identification", "A"
    positions.Add "status", "B"
    positions.Add "nom_complet", "C"
    positions.Add "derniere_maj", "D"
    positions.Add "type_materiel", "E"
    positions.Add "marque", "F"
    positions.Add "modele", "G"
    positions.Add "os", "H"
    positions.Add "no_serie", "I"
    positions.Add "date_crea_fiche_teck", "J"
    positions.Add "cree_par", "K"
    positions.Add "donateur", "L"
    positions.Add "coord_donateur", "M"
    positions.Add "date_recpt_don", "N"
    positions.Add "comment_don", "O"
    positions.Add "cpu", "P"
    positions.Add "cpu_indice", "Q"
    positions.Add "mem", "R"
    positions.Add "type_disk", "S"
    positions.Add "size_disk", "T"
    positions.Add "note_audit", "U"
    positions.Add "pond_teck", "V"
    positions.Add "pond_aspect", "W"
    positions.Add "total", "X"
    positions.Add "cat_vente_calculee", "Y"
    positions.Add "prix_vente_calculee", "Z"
    positions.Add "comm_prix", "AA"
    positions.Add "benevole_charge_recond", "AB"
    positions.Add "recond_hors_pa", "AC"
    positions.Add "date_pa", "AD"
    positions.Add "comm_recond", "AE"
    positions.Add "utilisation_cle_win_10", "AF"
    positions.Add "ctrl_connect", "AG"
    positions.Add "eff_donnees", "AH"
    positions.Add "dispo_vente_se", "AI"
    positions.Add "date_vente", "AJ"
    positions.Add "comm_vente", "AK"
    positions.Add "taille_ecran", "AL"
    positions.Add "alim_chargeur", "AM"
    positions.Add "cap_res_bat", "AN"
    positions.Add "auto_bat", "AO"
    positions.Add "bluetooth", "AP"
    positions.Add "webcam", "AQ"
    positions.Add "clavier_pc_fixe", "AR"
    positions.Add "saccoche", "AS"
    positions.Add "souris", "AT"
    positions.Add "dvd", "AU"
    positions.Add "lecteurSD", "AV"
    Set getPositionsIndex = positions
End Function

' Create the line of title (under a map form)
Function getTitlesMap()
    Dim titles
    Set titles = CreateObject("Scripting.Dictionary")
    ' Suivit
    titles.Add positions("identification"), "Identification"
    titles.Add positions("status"), "Statut"
    titles.Add positions("nom_complet"), "Nom complet"
    titles.Add positions("derniere_maj"), "Date de dernière maj"
    ' Matériel
    titles.Add positions("type_materiel"), "Type matériel"
    titles.Add positions("marque"), "Marque"
    titles.Add positions("modele"), "Modèle"
    titles.Add positions("os"), "Système d'exploitation"
    titles.Add positions("no_serie"), "Numéro de série"	
    ' DON
    titles.Add positions("date_crea_fiche_teck"), "Date de création fiche technique"
    titles.Add positions("cree_par"), "Créée par"
    titles.Add positions("donateur"), "Donateur"
    titles.Add positions("coord_donateur"), "Coordonnées Contact Donateur"
    titles.Add positions("date_recpt_don"), "Date de réception du don"
    titles.Add positions("comment_don"), "Commentaire sur le don"
    ' Catégorisation et calcul du prix de vente
    titles.Add positions("cpu"), "Processeur complet"
    titles.Add positions("cpu_indice"), "Indice CPU"
    titles.Add positions("mem"), "Mémoire (en Go)"
    titles.Add positions("type_disk"), "Type disque"
    titles.Add positions("size_disk"), "Taille disque (en Go)"
    titles.Add positions("note_audit"), "Note Audit"
    titles.Add positions("pond_teck"), "Pondération +/- technique"
    titles.Add positions("pond_aspect"), "Pondération +/- Aspect"
    titles.Add positions("total"), "Total"
    titles.Add positions("cat_vente_calculee"), "Catégorie de vente calculée"
    titles.Add positions("prix_vente_calculee"), "Prix de vente calculé"
    titles.Add positions("comm_prix"), "Commentaire Prix"
    ' suivit reconditionnement
    titles.Add positions("benevole_charge_recond"), "Bénévole en charge du reconditionnement"
    titles.Add positions("recond_hors_pa"), "Reconditionnement hors PA"
    titles.Add positions("date_pa"), "Date de prise en charge"
    titles.Add positions("comm_recond"), "Commentaire sur le recond."
    titles.Add positions("utilisation_cle_win_10"), "Utilisation d'une clé de licence Windows 10"
    titles.Add positions("ctrl_connect"), "Contrôle Connectique"
    titles.Add positions("eff_donnees"), "Effacement des données"
    ' vente
    titles.Add positions("dispo_vente_se"), "Disponible à la vente SE (O/N)"
    titles.Add positions("date_vente"), "Date de la vente"
    titles.Add positions("comm_vente"), "Commentaire vente"
    ' fiche technique
    titles.Add positions("taille_ecran"), "Taille écran"
    titles.Add positions("alim_chargeur"), "Alim. chargeur"
    titles.Add positions("cap_res_bat"), "Capacité résiduelle Batterie (%)"
    titles.Add positions("auto_bat"), "Autonomie Batterie (Min)"
    titles.Add positions("bluetooth"), "Bluetooth (Pres Abs KO)"
    titles.Add positions("webcam"), "Webcam (Pres Abs KO)"
    titles.Add positions("clavier_pc_fixe"), "Clavier PC fixe (Pres Abs KO)"
    titles.Add positions("saccoche"), "Sacoche (O/N)"
    titles.Add positions("souris"), "Souris (Pres Abs KO)"
    titles.Add positions("dvd"), "DVD / GRAVEUR (Pres Abs KO)"
    titles.Add positions("lecteurSD"), "Lecteur SD (Pres Abs KO)"
    Set getTitlesMap = titles
End Function


' Create a row at line in given sheet for an input object of variable form
Function sheetCreateRow(sheet, line, obj) 
	if ( Typename(obj) = "Dictionary" ) THEN
		sheetCreateRowFromHashMap sheet, line, obj
	elseif ( IsArray(obj) ) THEN
		sheetCreateRowFromArray sheet, line, obj
	ELSE
		MsgBox("obj " & Typename(obj) & " not supported")
		wscript.Quit
	END IF
End Function

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

' Get absolute path from a relative one
Function getAbsoluteFilenameFromRelative(name) 
	Set fso =  CreateObject("Scripting.FileSystemObject")
	getAbsoluteFilenameFromRelative = fso.GetAbsolutePathName(name)
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

' Create the sheet or open if existing
Function sheetOpenOrCreate(fname)
	fname = getAbsoluteFilenameFromRelative(fname)
	Set fso =  CreateObject("Scripting.FileSystemObject")
	if ( fso.FileExists(fname) ) then
		Set sheetOpenOrCreate = openExisting(fname)
	else
		Set sheetOpenOrCreate = sheetCreateInital()
	end if
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

' Create a entry in the sheet from the data found on this computer
Function sheetEntryFromThisPC(sheet, line)
    Dim thisPC
    Set thisPC = CreateObject("Scripting.Dictionary")
    thisPC.Add positions("cpu"), getCPU()
    thisPC.Add positions("mem"), getInstalledRAMgo()
    thisPC.Add positions("size_disk"), Round(getDiskSpaceGo(), 2)

    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)
    For Each objItem in colItems
       thisPC.Add positions("marque"), objItem.Vendor
       thisPC.Add positions("modele"), objItem.Version
    Next
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
    For Each objItem in colItems
	thisPC.Add positions("os"), objItem.Caption & "(" & objItem.Version & ")"
        thisPC.Add positions("no_serie"), objItem.SerialNumber
    Next

    Set objSession = CreateObject("Microsoft.Update.Session")
    Set objSearcher = objSession.CreateUpdateSearcher
    Set colHistory = objSearcher.QueryHistory(1, 1)
    For Each objEntry in colHistory
       thisPC.Add positions("derniere_maj"), objEntry.Date
    Next

    thisPC.Add positions("bluetooth"), bluetoothSupported()
    thisPC.Add positions("taille_ecran"), getScreenResolutionPx()
    thisPC.Add positions("date_crea_fiche_teck"), curDate()
    thisPC.Add positions("nom_complet"), getNomComplet()
    sheetCreateRow sheet, line, thisPC
End Function

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

' Same as sheetNewEntryFromThisPC but update entry if already in (by serial number)
Function sheetUpdateOrNewEntryFromThisPC(sheet)
	Dim line
	line = sheetThisPCinSheet(sheet)
	if line=-1 then
		sheetNewEntryFromThisPC(sheet)
	else
		sheetEntryFromThisPC sheet, line
	end if
End Function

' Create a new entry in the sheet from the data found on this computer
Function sheetNewEntryFromThisPC(sheet)
	Dim line
	line = usedRows(sheet, positions("cpu")) + 1
	sheetNewEntryFromThisPC = sheetEntryFromThisPC(sheet, line)
End Function
