' Generic function to create reports this lib must be used with a particular implementation (eg: excel)
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

' Get title of the sheet
Function getSheetTitle()
	getSheetTitle = "Tous les reconditionnements"
End Function

' Get subject of the sheet
Function getSheetSubject()
	getSheetSubject = "Reconditionnements"
End Function

' Get author of the sheet
Function getSheetAuthor()
	getSheetAuthor = "Emmaüs"
End Function

' Get titles of the very first line
' excel : RGB()
'https://docs.microsoft.com/fr-fr/office/vba/api/excel.font.color
'https://docs.microsoft.com/en-us/office/vba/api/excel.interior.color
' open office : hexa rgb
' Blue-Green-Red triplet
Function getBigTitles()
	Set titles = CreateObject("Scripting.Dictionary")
	titles.Add "suivi", CreateObject("Scripting.Dictionary")
	titles("suivi").Add "text", "SUIVI"
	titles("suivi").Add "text.color", &HFFFFFF
	titles("suivi").Add "bg", &H7e8187
	titles.Add "material", CreateObject("Scripting.Dictionary")
	titles("material").Add "text", "MATERIEL"
	titles("material").Add "text.color", &HFFFFFF
	titles("material").Add "bg", &H0055ff
	titles.Add "don", CreateObject("Scripting.Dictionary")
	titles("don").Add "text", "DON"
	titles("don").Add "text.color", &HFFFFFF
	titles("don").Add "bg", &H005c05
	titles.Add "cat", CreateObject("Scripting.Dictionary")
	titles("cat").Add "text", "CATEGORISATION ET CALCUL DU PRIX DE VENTE"
	titles("cat").Add "text.color", &HFFFFFF
	titles("cat").Add "bg", &Hdb852e
	titles.Add "suivi_recon", CreateObject("Scripting.Dictionary")
	titles("suivi_recon").Add "text", "SUIVI DU RECONDITIONNEMENT"
	titles("suivi_recon").Add "text.color", &HFFFFFF
	titles("suivi_recon").Add "bg", &H2a039e
	titles.Add "vente", CreateObject("Scripting.Dictionary")
	titles("vente").Add "text", "VENTE"
	titles("vente").Add "text.color", &H000000
	titles("vente").Add "bg", &H5de381
	titles.Add "teck", CreateObject("Scripting.Dictionary")
	titles("teck").Add "text", "FICHE TECHNIQUE"
	titles("teck").Add "text.color", &H000000
	titles("teck").Add "bg", &Hab521b
	Set getBigTitles = titles
End Function

' Apply inspectorRGBToLongRgb to all colors in titles
Function convertBigTitlesToLongRgb(titles)
	FOR EACH k IN titles
		titles(k)("bg") = inspectorRGBToLongRgb(titles(k)("bg"))
		titles(k)("text.color") = inspectorRGBToLongRgb(titles(k)("text.color"))
	NEXT
	Set convertBigTitlesToLongRgb = titles
End Function

' Convert from hex rgb to vbscript rgb
' https://www.oreilly.com/library/view/vbscript-in-a/1565927206/re159.html#:~:text=The%20RGB%20color%20value%20represents,greater%20than%20255%20is%20255.
Function inspectorRGBToLongRgb(color)
	Dim r, g, b
	r = RightShift((color And ((2^24)-1)), 16)
	g = RightShift((color And ((2^16)-1)), 8)
	b = (color And ((2^8)-1))
	inspectorRGBToLongRgb = RGB(r, g, b)
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

' Get absolute path from a relative one
Function getAbsoluteFilenameFromRelative(name) 
	Set fso =  CreateObject("Scripting.FileSystemObject")
	getAbsoluteFilenameFromRelative = fso.GetAbsolutePathName(name)
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

' Create a entry in the sheet from the data found on this computer
Function sheetEntryFromThisPC(sheet, line)
    Dim thisPC
    Set thisPC = CreateObject("Scripting.Dictionary")
	thisPC.Add positions("clavier_pc_fixe"), getKeyboard()
	thisPC.Add positions("alim_chargeur"), guessAlimChargeur()
    thisPC.Add positions("cpu"), getCPU()
	thisPC.Add positions("cpu_indice"), getCPUindice()
	thisPC.Add positions("type_materiel"), getMaterielType()
    thisPC.Add positions("mem"), getInstalledRAMgo()
    thisPC.Add positions("size_disk"), Round(getDiskSpaceGo(), 2)
	thisPC.Add positions("type_disk"), getDiskType()
	thisPC.Add positions("dvd"), getCDROMinfos()
	thisPC.Add positions("cap_res_bat"), getBatteryCapResid()
	thisPC.add positions("auto_bat"), getBatteryAmountTimeExpected()
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




' Same as sheetNewEntryFromThisPC but update entry if already in (by serial number)
Function sheetUpdateOrNewEntryFromThisPC(sheet)
	Dim line
	line = sheetThisPCinSheet(sheet)
	if line=-1 then
		sheetNewEntryFromThisPC sheet
	else
		sheetEntryFromThisPC sheet, line
	end if
	IF TypeName(sheet)="Worksheet" Or TypeName(sheet)="ISheet" THEN
		Set sheetUpdateOrNewEntryFromThisPC = sheet
	ELSE
		sheetUpdateOrNewEntryFromThisPC = sheet
	END IF
End Function

' Create a new entry in the sheet from the data found on this computer
Function sheetNewEntryFromThisPC(sheet)
	Dim line
	line = usedRows(sheet, positions("cpu")) + 1
	sheetNewEntryFromThisPC = sheetEntryFromThisPC(sheet, line)
End Function

' Get a compatible output format from a given name
Function getCompatOutputFmt(fname, destExt)	
	getCompatOutputFmt = reReplace(fname, "\.[^\.]+$", "") & destExt ' remove extra extension in filename
End Function

' get avaliable format for this computer
Function getAvaliableFormats()
	ReDim fmts(0)
	fmts(0)="csv"
	On Error Resume Next
	Err.Clear
	Set myObj = CreateObject("Excel.Application")
	If Err.Number <> 0 Then
	Else
		ReDim Preserve fmts(1)
		fmts(1)="xlsx"
	End If
	Err.Clear
	Set myObj = CreateObject("com.sun.star.ServiceManager")
	If Err.Number <> 0 Then
	Else
		ReDim Preserve fmts(UBound(fmts)+1)
		fmts(UBound(fmts)) = "ods"
	End If
	Err.Clear
	Set myObj = CreateObject("PlanMaker.Application")
	If Err.Number <> 0 Then
	Else
		ReDim Preserve fmts(UBound(fmts)+1)
		fmts(UBound(fmts)) = "pmdx"
	End If
	 getAvaliableFormats = fmts
End Function

' Get only ext name from the complete filename
Function onlyExtName(filename)
	On Error Resume Next
	Err.Clear
	Dim matches
	Set oRegExp = New RegExp
	oRegExp.Pattern = "\.([^\.]+)$"
	oRegExp.Global = True
	Set matches = oRegExp.Execute(filename)
	onlyExtName = reReplace(matches(0), "^\.", "")
	If Err.Number <> 0 Then
		onlyExtName = ""
	end if
	On Error Goto 0
End Function

' Get filename with extension compatible with this lib
Function getOutputFile(fname)
	Dim ext 
	ext = onlyExtName(fname)
	if ext="" then
		ext = getPreferredExtension()
	end if
	getOutputFile = getCompatOutputFmt(fname, "." & ext)
End Function

' load backend for a given extension
Function loadBackendForExtension(extname)
	extname = onlyExtName(extname)
	Set lst = CreateObject("Scripting.Dictionary")
	lst.Add "csv", "libCsvReports.vbs"
	lst.Add "ods", "libOpenOfficeReports.vbs"
	lst.Add "xlsx", "libExcelReports.vbs"
	lst.Add "pmdx", "libPlanMakerReports.vbs"
	load(lst(extname))
End Function



' Load prefered backend for this computer
Function loadPreferedBackend()
	' Load a backend from what is avaliable
	On Error Resume Next
	Err.Clear
	Set myObj = CreateObject("Excel.Application")
	If Err.Number <> 0 Then
		Err.Clear
		Set myObj = CreateObject("com.sun.star.ServiceManager")
		If Err.Number <> 0 Then
			Err.Clear
			Set myObj = CreateObject("PlanMaker.Application")
			If Err.Number <> 0 Then
				load("libCsvReports.vbs")
			Else
				load("libPlanMakerReports.vbs")
			End If
		Else
			load("libOpenOfficeReports.vbs")
		End If
	    	Err.Clear
	Else
	    load("libExcelReports.vbs")
	End If
	On Error Goto 0
End Function
