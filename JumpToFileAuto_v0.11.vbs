'******************************************************************************************************
'*** Script Name:		 Jump to file
'*** Script Description: With this script you can search within nowplaying list, 
'***					 play selected song and you can make your on queue list
'***
'*** Original Author:	 Danijel Sokač
'***
'*** Contact:			 dsokac1@gmail.com
'***
'*** Disclaimer:		 This software is provided 'as-is', without any express or implied warranty.
'*** 			 		 In no event will the author be held liable for any damages arising from the
'*** 			 		 use of this software.
'******************************************************************************************************

Option Explicit
	'songIndex 				- index pjesme koja se treba pustiti
	'LB 					- za pristup listboxu
	'objSearchSongList 		- lista svih pjesama iz nowplaying
	'objSearchSongList2 	- lista filtriranih pjesama 
	'max 					- maksimalan broj ispisanih pjesama
	'ex_tekst				- pamtni prethodni unos u textbox
	'Form					- objekt forme koja se prikazuje na ekranu
	'mode					- 1 - Jump to file
	'						  2 - Queue list
	'queueList				- lista pjesama koje su u stanju čekanja
	'queuePlaylist			- sadrži queue playlistu 
	'menuAddToQueueHotkey 	- Prečac za AddToQueue
	'menuQueueListHotkey	- Prečac za prikaz queue liste
	'menuJTFHotkey			- Prečac za otvaranje Jump to file
	'objSongList			- lista pjesama iz queue liste
	'exQueuedIndex			- Index pjesme i queue liste koja je zadnja svirana
dim songIndex,  LB, objSearchSongList, objSearchSongList2, Form, queueList, queuePlaylist
dim max : max = 100
dim maxSearchIndex : maxSearchIndex = 0
dim ex_tekst : ex_tekst = ""
dim mode : mode = 1
dim menuAddToQueueHotkey : menuAddToQueueHotkey = "Alt+q"
dim menuQueueListHotkey : menuQueueListHotkey = "q"
dim menuJTFHotkey : menuJTFHotkey = "j"
dim objSongList
dim exQueuedIndex : exQueuedIndex = -1
dim forcePlay : forcePlay = false

Sub OnStartup
	dim objMenuItem
	Set objSearchSongList = SDB.Player.CurrentSongList
	Set queuePlaylist = SDB.PlaylistByTitle("")
	Set queuePlaylist = queuePlaylist.CreateChildPlaylist("QueuePlaylist")
		'Add to queue list menu
	Set objMenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_Pop_NP, 2, 1)
	objMenuItem.Caption = "Add to &queue list"
	objMenuItem.Shortcut = menuAddToQueueHotkey
	objMenuItem.UseScript = Script.ScriptPath
	objMenuItem.OnClickFunc = "OnAddToQueueMenuClicked"
	objMenuItem.Visible = True
	
	Set objMenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_Pop_NP_MainWindow, 2, 1)
	objMenuItem.Caption = "Add to &queue list"
	objMenuItem.Shortcut = menuAddToQueueHotkey
	objMenuItem.UseScript = Script.ScriptPath
	objMenuItem.OnClickFunc = "OnAddToQueueMenuClicked"
	objMenuItem.Visible = True
	
		'Open Jump to file
	Set objMenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_Play, 2, 0)
	objMenuItem.Caption = "&Jump to file"
	objMenuItem.Shortcut = menuJTFHotkey
	objMenuItem.UseScript = Script.ScriptPath
	objMenuItem.OnClickFunc = "OnJTFMenuClicked"
	objMenuItem.Visible = True
		'Open queue list
	Set objMenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_Play, 2, 0)
	objMenuItem.Caption = "Open queue list"
	objMenuItem.Shortcut = menuQueueListHotkey
	objMenuItem.UseScript = Script.ScriptPath
	objMenuItem.OnClickFunc = "OnQueueListMenuClicked"
	objMenuItem.Visible = True
	
	Script.RegisterEvent SDB, "OnPlay", "NextTrack"
End Sub


Sub NextTrack
	'SDB.MessageBox "Next Track!", mtInformation, Array(mbOk)
	Set objSongList = queuePlaylist.Tracks
		'Ako ima pjesama u queue listi pusti iz nje
	If objSongList.Count AND SDB.Player.CurrentSongIndex <> exQueuedIndex AND NOT forcePlay Then
		Set objSearchSongList = SDB.Player.CurrentSongList
		PlayQueuedSong
		exQueuedIndex = FindSongIndex(objSongList.Item(0))
		queuePlaylist.RemoveTrackNoConfirmation objSongList.Item(0)
		Set objSongList = queuePlaylist.Tracks
		If objSongList.Count = 0 Then
			exQueuedIndex = -1
		End if
		BrisiLB
		if mode = 1 Then
			ListSongs
		Else
			ListQueuedSongs
		End if
	Else if forcePlay Then
		forcePlay = False
		End if
	End if
End Sub

Sub PlayQueuedSong
		'objSongData		- objekt tipa SongData
	dim objSongData, shuffleFlag
	Set objSongData = objSongList.Item(0)
		'Preko ID-a pronađi index pjesme u NowPlaying listi
	songIndex = FindSongIndex(objSongData)
	SDB.Player.CurrentSongIndex = songIndex
	if not SDB.Player.isPlaying Then
		SDB.Player.Play 
	End if
End Sub

Sub OnAddToQueueMenuClicked(objMenuItem)
	dim objSongData, objSelectedSong
	Set objSelectedSong = SDB.SelectedSongList
	Set objSongData = objSelectedSong.Item(0)
	AddToQueueList objSongData
End Sub

Sub OnQueueListMenuClicked(objMenuItem)
	mode = 2
	JumpToFile
End Sub

Sub OnJTFMenuClicked(objMenuItem)
	mode = 1
	JumpToFile
End Sub

Sub AddToQueueList (objSongData)
	dim i, objTmpSongData
	if objSongData is Nothing Then
		Exit Sub
	End if
	Set queueList = queuePlaylist.Tracks
	for i = 0 to queueList.Count - 1
		Set objTmpSongData = queueList.Item(i)
		if objSongData.SongID = objTmpSongData.SongID Then
			Exit for
		End if
	Next
	if i = queueList.Count Then
		queuePlaylist.AddTrack objSongData
	else
		queuePlaylist.RemoveTrackNoConfirmation queueList.Item(i)
	End if
	Set queueList = queuePlaylist.Tracks
End Sub

	'This subroutine pop up JumpToFile box
Sub JumpToFile
	Set objSearchSongList = SDB.Player.CurrentSongList
	Set queueList = queuePlaylist.Tracks
	IzradiFormu
End Sub

	'Opisuje izgled prozora ovisno o mode-u
Sub IzradiFormu
	dim textbox, ButtonJTF, ButtonQF, ButtonClose, ButtonCM, ButtonUp, ButtonDown, ButtonRemove, i
	dim ButtonMoveAfterCurrent, ButtonRemoveAll
		
		'Postavlja glavni okvir prozora
	Set Form = SDB.UI.NewForm
	Form.Common.SetClientRect 100, 100, 500, 440			'veličina prozora
	Form.BorderStyle = 2									'standard resizable window
	Form.FormPosition = 4									'prikaži u sredini ekrana
	Form.StayOnTop = True
	
		'list box - prostor za ispis pjesama
	Set LB = SDB.UI.NewListBox(Form)
	LB.Common.SetRect 5, 30, 485, 360						'veličina polja za pjesme
	Script.RegisterEvent LB.Common, "OnDblClick", "PlaySongDblClick"
		'ChangeMod Button
	Set ButtonCM = SDB.UI.NewButton(Form)
	ButtonCM.Common.SetClientRect 5, 390, 480, 20
	Script.RegisterEvent ButtonCM, "OnClick", "ChangeMode"
	
	if mode = 1 Then
		Form.Caption = "Jump To File"						'Ime prozora
			'textbox
		Set textbox = SDB.UI.NewEdit(Form)
		textbox.Common.SetClientRect 5, 5, 480, 50			'veličina textboxa
		Script.RegisterEvent textbox, "OnChange", "search"	'registriraj promjenu sadržaja i pozovi search
			'ispiši početni popis pjesama
		ListSongs
			'JumpToFile Button
		Set ButtonJTF = SDB.UI.NewButton(Form)
		ButtonJTF.Common.SetClientRect 5, 415, 100, 20			'veličina tipke
		ButtonJTF.Caption = "&Jump to file"						'naziv tipke
		ButtonJTF.Default = True								'Pritiskom na enter aktivira se tipka
		Script.RegisterEvent ButtonJTF, "OnClick", "PlaySong"	'pozovi playsong kod pritiska tipke
			'QueueFile Button
		Set ButtonQF = SDB.UI.NewButton(Form)
		ButtonQF.Common.SetClientRect 110, 415, 100, 20
		ButtonQF.Caption = "&Queue file"
		Script.RegisterEvent ButtonQF, "OnClick", "OnAddToQueueClicked"	'Pozovi QueueSong kad se pritisne tipka
		
			'MoveAfterCurrent Button
		Set ButtonMoveAfterCurrent = SDB.UI.NewButton(Form)
		ButtonMoveAfterCurrent.Common.SetClientRect 215, 415, 150, 20		'veličina tipke
		ButtonMoveAfterCurrent.Caption = "Move &after current"				'naziv tipke
		Script.RegisterEvent ButtonMoveAfterCurrent, "OnClick", "MoveAfterCurrent"	'pozovi playsong kod pritiska tipke
		
			'ChangeMode button name
		ButtonCM.Caption = "&To Queue list"
	else
		Form.Caption = "Queue list"							'Ime prozora
			'ChangeMode button name
		ButtonCM.Caption = "&To Jump to file"
			'ButtonUp (move song up in playlist)
		Set ButtonUp = SDB.UI.NewButton(Form)
		ButtonUp.Common.SetClientRect 5, 415, 70, 20
		ButtonUp.Caption = "&Up"
		Script.RegisterEvent ButtonUp, "OnClick", "MoveUp"
			'ButtonDown (move song down in playlist)
		Set ButtonDown = SDB.UI.NewButton(Form)
		ButtonDown.Common.SetClientRect 80, 415, 70, 20
		ButtonDown.Caption = "&Down"
		Script.RegisterEvent ButtonDown, "OnClick", "MoveDown"
			'ButtonRemove (Remove track from playlist)
		Set ButtonRemove = SDB.UI.NewButton(Form)
		ButtonRemove.Common.SetClientRect 155, 415, 100, 20
		ButtonRemove.Caption = "&Remove"
		Script.RegisterEvent ButtonRemove, "OnClick", "Remove"
			'ButtonRemoveAll (Remove track from playlist)
		Set ButtonRemoveAll = SDB.UI.NewButton(Form)
		ButtonRemoveAll.Common.SetClientRect 260, 415, 100, 20
		ButtonRemoveAll.Caption = "Remove &all"
		Script.RegisterEvent ButtonRemoveAll, "OnClick", "RemoveAll"
		
			'List queued songs
		BrisiLB
		ListQueuedSongs
	End if
	
		'Close button
	Set ButtonClose = SDB.UI.NewButton(Form)
	ButtonClose.Common.SetClientRect 380, 415, 100, 20
	ButtonClose.Caption = "&Close"
	Script.RegisterEvent ButtonClose, "OnClick", "CloseForm"
	ButtonClose.Cancel = True

	Form.Common.Visible = True					
	SDB.Objects("Form") = Form
End Sub

	'Zatvori formu
Sub CloseForm
	Form.Common.Visible = False
	objSearchSongList2 = Empty
End Sub

	'Ispiši pjesme ovisno o trenutnoj koja svira
Sub ListSongs
	dim objSongData, objSongList, currentSongIndex, i
	Set objSongList = SDB.Player.CurrentSongList
		'Index pjesme koja trenutno svira
	currentSongIndex = SDB.Player.CurrentSongIndex
		'obriši prozor za pjesme
	brisiLB
		'Ako postoji pjesama u search listi
	If isObject(objSearchSongList2) Then
		for i = 0 to objSearchSongList2.Count - 1
				ispisi objSearchSongList2, i
		Next
		'ako ima u nowplaying listi više od max pjesama ispiši samo max pjesama
	Elseif objSongList.Count > max Then
				'ako stanu sve pjesme od trenutna +- max/2 ispiši ih
		If (currentSongIndex > max\2) And (currentSongIndex < (objSongList.Count - max\2 - 1)) Then
			for i = currentSongIndex - max\2 to currentSongIndex + max\2
				ispisi objSongList, i
			Next
				'ako je index trenutne manji od max/2 ispiši prvih max
		Elseif currentSongIndex <= max\2 Then
			for i = 0 to max
				ispisi objSongList, i
			Next
			'inače ispiši zadnjih max
		Else
			for i = objSongList.Count - max to objSongList.Count - 1
				ispisi objSongList, i
			Next
		End If
		'ako ima manje od max pjesama ispiši sve
	Else
		for i = 0 to objSongList.Count - 1
			ispisi objSongList, i
		Next
	End If
End Sub

	'Ispiši pjesme iz queue liste
Sub ListQueuedSongs
	dim i, objSongData, StringTitle, StringArtist
	queueList = Empty
	Set queueList = queuePlaylist.Tracks
	for i = 0 to queueList.Count - 1
		Set objSongData = queueList.Item(i)
		StringTitle = objSongData.Title
		StringArtist = objSongData.ArtistName
		LB.Items.Add (i + 1) & ". " & StringArtist & " - " & Stringtitle
	Next		
End Sub

	'briše sve prikazane pjesme iz LB-a
Sub brisiLB()
	dim i: i = 0
	do while i < LB.Items.Count
		LB.Items.Delete i
	Loop
End Sub

	'Ispisuje ime pjevača i naziv pjesme iz liste objSongList sa indexsom i
Sub ispisi(objSongList, i)
	dim objSongData, StringTitle, StringArtist, j, tmp_objSongData
	Set objSongData = objSongList.Item(i)
	Set queueList = queuePlaylist.Tracks
	StringTitle = objSongData.Title
	StringArtist = objSongData.ArtistName
	for j = 0 to queueList.Count - 1
		Set tmp_objSongData = queueList.Item(j)
		'SDB.MessageBox "For, j = " & j, mtInformation, Array(mbOk)
		if  tmp_objSongData.SongID = objSongData.SongID Then
			'SDB.MessageBox "Pišem", mtInformation, Array(mbOk)
			LB.Items.Add StringArtist + " - " + Stringtitle + "    [" + CStr(j+1) + "]"
			j = -1
			Exit for
		End if
	Next
	if j <> -1 Then
		LB.Items.Add StringArtist + " - " + Stringtitle
	End if
End Sub

Sub PlaySongDblClick(control)
	if mode = 1 Then
		PlaySong
	End if
End Sub

	'Pronađi index označene pjesme
Function FindSongIndex(objSongData)
	dim i, tmp_objSongData
	for i = 0 to objSearchSongList.Count - 1
		Set tmp_objSongData = objSearchSongList.Item(i)
		if tmp_objSongData.Path = objSongData.Path Then
			FindSongIndex = i
			Exit For
		End if
	Next
End Function

	'Pušta pjesmu koja je označena u LB
Sub PlaySong
		'objSongData		- objekt tipa SongData
	forcePlay = True
	dim objSongData
	Set objSongData = GetSelectedSongData
	if objSongData is Nothing Then
		Exit Sub
	End if
	
		'Preko ID-a pronađi index pjesme u NowPlaying listi
	songIndex = FindSongIndex(objSongData)
	SDB.Player.CurrentSongIndex = songIndex
	if not SDB.Player.isPlaying Then
		SDB.Player.Play 
	End if
	if isObject(objSearchSongList2) Then
		objSearchSongList2 = Empty
	End if
	CloseForm
End Sub

	'Radi pretragu nowplaying liste na temelju unosa u textbox
	'[in] control - textbox objekt
Sub search(control)
		'objRE 				- Regexp objekt za uspoređivanje stringova
		'objSongData		- objekt za čitanje podatak o pjesmi
		'i,j				- brojač petlje
		'tmp_SearchSongList - privremena song lista za pohranu rezultata pretrage
		'patternList		- lista s riječima unesenog teksta
		'flag				- zastavica za provjeru je li petlja dosla do kraja
	dim objRE, objSongData, i, tmp_SearchSongList, patternList, j, flag
	flag = True
	Set objRE = New RegExp
	objRE.IgnoreCase = True						'Case unsensitive
	objRE.Pattern = ex_tekst					'Kao traženi pojam uzmi tekst iz textboxa
	patternList = Split(control.text, " ")
	brisiLB										'Briši pjesme iz LB-a
	Set tmp_SearchSongList = SDB.NewSongList	'Napravi privremenu song listu
	
		'Ako nije unesen tekst izlistaj max pjesama oko trenutne i izađi iz procedure
	if control.text = "" Then
		maxSearchIndex = 0
		'SDB.MessageBox "Tekst: " & control.text & ", index: " & maxSearchIndex, mtInformation, Array(mbOk)
		objSearchSongList2 = Empty
		ListSongs
		Exit Sub
		'Prvo prođi kroz dosad filtriranu listu ako postoji i ako nije brisan znak iz prezhodnog unosa
	Elseif not objRE.Test(control.text) Then
		'SDB.MessageBox "Info", mtInformation, Array(mbOk)
		maxSearchIndex = 0
		if isObject(objSearchSongList2) Then
			objSearchSongList2 = Empty
		End if
	Elseif isObject(objSearchSongList2) Then
		'SDB.MessageBox "Lista2.Count = " & objsearchSongList2.Count, mtInformation, Array(mbOk)
		for i = 0 to objSearchSongList2.Count - 1
			Set objSongData = objSearchSongList2.Item(i)		'pročitaj pjesmu iz filtrirane liste
			for each j in patternList
				'SDB.MessageBox "Tekst: " & j & ", index: " & maxSearchIndex & "\n i = " & i, mtInformation, Array(mbOk)
				objRE.Pattern = j								'Kao traženi pojam uzmi riječ iz textboxa
					'usporedi objRE.pattern i podatke od pjesme
				If NOT objRE.Test(objSongData.ArtistName) AND NOT objRE.Test(objSongData.Title)	AND NOT objRE.Test(objSongData.AlbumArtistName) AND NOT objRE.Test(objSongData.Year) AND NOT objRE.Test(objSongData.AlbumName) AND NOT objRE.Test(objSongData.Path) Then
					flag = False
					Exit For
				End If
			Next
			if flag Then
				'SDB.MessageBox "Pjesma: " & objSongData.Title & "-" & objSongData.ArtistName, mtInformation, Array(mbOk)
				tmp_SearchSongList.Add objSongData
				ispisi objSearchSongList2, i					'Ispiši pjesmu koja odgovara traženom pojmu
				'objSearchSongList2.Delete i
			End if
			flag = True
			'SDB.MessageBox "Wait" & i, mtInformation, Array(mbOk)
		Next
		objSearchSongList2 = Empty
	End if
		'Traži pjesme iz nowplaying liste (od zadnje koja je provjeravana) koje sadrže traženi pojam 
	i = maxSearchIndex
	'SDB.MessageBox "i = " & i & "Index = " & maxSearchIndex & " tmp.Count = " & tmp_SearchSongList.Count, mtInformation, Array(mbOk)
	Do while i < objSearchSongList.Count AND tmp_SearchSongList.Count < max
		'SDB.MessageBox "i = " & i & " tmp.Count = " & tmp_SearchSongList.Count, mtInformation, Array(mbOk)
		Set objSongData = objSearchSongList.Item(i)		'pročitaj pjesmu iz nowplaying liste
		for each j in patternList
			objRE.Pattern = j							'Kao traženi pojam uzmi riječ iz textboxa
				'usporedi objRE.pattern i podatke od pjesme
			If NOT objRE.Test(objSongData.ArtistName) AND NOT objRE.Test(objSongData.Title) AND NOT objRE.Test(objSongData.AlbumArtistName) AND NOT objRE.Test(objSongData.Year) AND NOT objRE.Test(objSongData.AlbumName) AND NOT objRE.Test(objSongData.Path) Then
				flag = False
				Exit For
			End If
		Next
		if flag Then
			tmp_SearchSongList.Add objSongData
			ispisi objSearchSongList, i					'Ispiši pjesmu koja odgovara traženom pojmu
			'SDB.MessageBox "tmp.count = " & tmp_SearchSongList.Count, mtInformation, Array(mbOk)
				'Ako smo pronašli više od max pjesama, zaustavi pretragu
			If tmp_SearchSongList.Count >= max OR  i = objSearchSongList.Count - 1 Then
				'SDB.MessageBox "Index = " & maxSearchIndex, mtInformation, Array(mbOk)
				maxSearchIndex = i + 1
				Exit Do
			End if
		elseif i = objSearchSongList.Count - 1 Then'ne provjerava zadnjeg!!!!!
			maxSearchIndex = i + 1
			Exit Do
		End if
		flag = True
		i = i + 1
	Loop
	ex_tekst = control.text
	Set objSearchSongList2 = tmp_SearchSongList				'Spremi privremenu listu u objSearchSongList2
End Sub

	'Promijeni mode: Jump to file <-> Queue list
Sub ChangeMode
	if mode = 1 Then
		mode = 2
	Else
		mode = 1
	End if
	CloseForm
	IzradiFormu
End Sub

	'Stavi označenu pjesmu iza trenutne koja svira i isključi shuffle
Sub MoveAfterCurrent
	dim currentSongIndex, objSongData
	currentSongIndex = SDB.Player.CurrentSongIndex
	Set objSongData = GetSelectedSongData
	if objSongData is Nothing Then
		Exit Sub
	End if
		'Preko ID-a pronađi index pjesme u NowPlaying listi
	songIndex = FindSongIndex(objSongData)
		'Stavi pjesmu koja treba svirati iza trenutne pjesme ovisno o njihovom indexu u listi
	if songIndex < currentSongIndex OR songIndex = currentSongIndex Then
		SDB.Player.PlaylistMoveTrack songIndex, currentSongIndex
	Else
		SDB.Player.PlaylistMoveTrack songIndex, currentSongIndex + 1
	End if
	SDB.Player.isShuffle = False
End Sub

	'Get song data of song selected in listbox
Function GetSelectedSongData
	dim currentSongIndex
	currentSongIndex = SDB.Player.CurrentSongIndex	
	If LB.ItemIndex = -1 Then
		Set GetSelectedSongData = Nothing
		Exit Function
	End If
	'Ako postoji pjesama u objSearcjSongList2 (imamo filtrirane pjesme) čitaj iz te liste
	if isObject(objSearchSongList2) Then
		Set GetSelectedSongData = objSearchSongList2.Item(LB.ItemIndex)
		'Ako su ispisane sve pjesme onda ovisno o indexu trenutne pjesme odredi koji se odmak koristi od početka playliste
		'kako bi znali koja je pjesma označena
	elseif currentSongIndex <= max\2 OR objSearchSongList.Count < max Then
		Set GetSelectedSongData = objSearchSongList.Item(LB.ItemIndex)
	elseif (currentSongIndex > max \2) And (currentSongIndex < (objSearchSongList.Count - max\2 - 1)) Then
		Set GetSelectedSongData = objSearchSongList.Item(LB.ItemIndex + currentSongIndex - max\2)
	else
		Set GetSelectedSongData = objSearchSongList.Item(LB.ItemIndex + objSearchSongList.Count - max)
	End if
End Function

	'Dodavanje u queue listu
Sub OnAddToQueueClicked
	dim objSongData
	Set queueList = queuePlaylist.Tracks
	Set objSongData = GetSelectedSongData
	AddToQueueList objSongData
	BrisiLB
	ListSongs
End Sub


	'Pomakni jednu poziciju gore u playlisti
Sub MoveUp
	dim objSongData, objBeforeSongData
	If LB.ItemIndex = -1 OR LB.ItemIndex = 0 Then
		Exit Sub
	End if
	Set objSongData = queueList.Item(LB.ItemIndex)
	Set objBeforeSongData = queueList.Item(LB.ItemIndex - 1)
	queuePlaylist.MoveTrack objSongdata, objBeforeSongData
	brisiLB
	ListQueuedSongs
End Sub


	'Pomakni jednu poziciju dolje u playlisti
Sub MoveDown
	dim objSongData, objBeforeSongData
	If LB.ItemIndex = -1 OR LB.ItemIndex = queueList.Count - 1 Then
		Exit Sub
	End if
	Set objSongData = queueList.Item(LB.ItemIndex)
	If LB.ITemIndex = queueList.Count - 2 Then
		Set objBeforeSongData = queueList.Item(LB.ItemIndex)
		Set objSongData = queueList.Item(LB.ItemIndex + 1)
		queuePlaylist.MoveTrack objSongdata, objBeforeSongData
	Else
		Set objBeforeSongData = queueList.Item(LB.ItemIndex + 2)
		queuePlaylist.MoveTrack objSongdata, objBeforeSongData
	End if
	brisiLB
	ListQueuedSongs
End Sub


	'Izbriši iz playliste
Sub Remove
	dim objSongList
	Set objSongList = queuePlaylist.Tracks
	If LB.ItemIndex <> -1 Then
		queuePlaylist.RemoveTrackNoConfirmation objSongList.Item(LB.ItemIndex)
		brisiLB
		ListQueuedSongs
		Set queueList = queuePlaylist.Tracks
		If queueList.Count = 0 Then
			exQueuedIndex = -1
		End if
	End if
End Sub

	'Izbriši sve pjesme iz queue liste
Sub RemoveAll
	dim objSongList, playlistCount, i
	Set objSongList = queuePlaylist.Tracks
	playlistCount = objSongList.Count
	For i = 0 to playlistCount - 1
		queuePlaylist.RemoveTrackNoConfirmation objSongList.Item(0)
		Set objSongList = queuePlaylist.Tracks
		playlistCount = objSongList.Count
	Next
	exQueuedIndex = -1
	brisiLB
	listQueuedSongs
	Set queueList = queuePlaylist.Tracks
End Sub

