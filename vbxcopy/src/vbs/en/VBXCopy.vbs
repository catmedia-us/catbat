'--------------------------------------------------------------------------------
' $Id: VBXCopy.vbs,v 1.9 2010/02/07 23:02:41 keilw Exp $
'--------------------------------------------------------------------------------

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' VBXCopy
'
' Version 1.0.12
'
' Copyright 1999-2016 Creative Arts & Technologies. All Rights reserved.
'
' Parts Copyright 1999-2000 Survey Computing. Alle Rechte vorbehalten.
' Parts Copyright 1998 Microsoft Corporation. Alle Rechte vorbehalten.
' 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' @todo: translate
' Informationen zur Codequalität:
'
' 1) Der folgende Code führt zahlreiche Zeichenfolgeänderungen 
'    durch, indem kurze Zeichenfolgen mit & (Operator) verkettet 
'    werden. Da diese Verkettung jedoch aufwändig ist, handelt es sich
'    um eine sehr ineffiziente Lösung zum Schreiben von Code.
'    Andererseits kann ein derartiger Code sehr gut gepflegt werden; er
'    wird hier verwendet, da dieses Programm zahlreiche Festplatten-
'    vorgänge durchführt und die Festplatte sehr viel langsamer
'    arbeitet als die für die Verkettung von Zeichenfolgen 
'    notwendigen Speichervorgänge. Beachten Sie, dass es sich um einen
'    Beispiel- und nicht um tatsächlichen Produktcode handelt.
'
' 2) Option Explicit wird verwendet, da der Zugriff auf deklarierte
'    Variablen schneller als der auf nicht deklarierte ist. Darüber
'    hinaus wird verhindert, dass sich Fehler in den Code einschleichen,
'    wie z. B. Rechtschreibfehler (DriveTypeCDORM an Stelle von
'    DriveTypeCDROM).
'
' 3) Fehlerbehandlungsmethoden wurde nicht integriert, um diesen Code
'    lesbarer zu gestalten. Es wurden zwar Vorkehrungen getroffen, dass
'    in allgemeinen Fällen keine Fehler auftreten, jedoch ist dies je
'    nach verwendetem Dateisystemen unterschiedlich. Verwenden Sie in
'    Ihrem Produktcode On Error Resume Next sowie das Error-Objekt,
'    um mögliche Fehler aufzufangen.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Global Variables
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const AppTitle = "VBXCopy"
Const AppVersion = "1.0.12"
Const AppLanguage = "en"

Dim gbDebug
Dim TabStop
Dim NewLine

Dim TestDrive
Dim TestFilePath
Dim TestTargetPath
Dim TestTargetDrive

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Drive.DriveType Values
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const DriveTypeRemovable = 1
Const DriveTypeFixed = 2
Const DriveTypeNetwork = 3
Const DriveTypeCDROM = 4
Const DriveTypeRAMDisk = 5

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' File.Attributes Values
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const FileAttrNormal  = 0
Const FileAttrReadOnly = 1
Const FileAttrHidden = 2
Const FileAttrSystem = 4
Const FileAttrVolume = 8
Const FileAttrDirectory = 16
Const FileAttrArchive = 32 
Const FileAttrAlias = 64
Const FileAttrCompressed = 128

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' File Open Constants
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PopUp Constants for Messages
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const PopUpButtonOK = 0 'Show [OK] button 
Const PopUpButtonOKCancel = 1 'Show [OK] and [Cancel] buttons 
Const PopUpButtonAbortRetryIgnore = 2 'Show [Abort], [Retry] and [Ignore] buttons 
Const PopUpButtonYesNoCancel = 3 'Show [Yes], [No] and [Cancel] buttons 
Const PopUpButtonYesNo = 4 'Show [Yes] and [No] buttons 
Const PopUpButtonRetryCancel = 5 'Show [Retry] and [Cancel] buttons 

Const PopUpIconStop = 16 'Show Stop Mark icon 
Const PopUpIconQuestion = 32 'Show Question Mark icon 
Const PopUpIconExclamation = 48 'Show Exclamation Mark icon 
Const PopUpIconInfo = 64 'Show Information Mark icon 


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Result and Error Constants
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const ShowResultNone = 0
Const ShowResultAll = 1 
Const ShowResultStart = 2
Const ShowResultFinal = 3
Const ShowResultErrors = 4
Const ShowResultInfo = 5

'Dim pars

Main()
Main()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' ShowDriveType
'
' Zweck: 
'
' Erstellt eine Zeichenfolge, die den Laufwerkstyp eines angegebenen
' Drive-Objekts beschreibt.
'
' Zeigt Folgendes 
'
' - Drive.DriveType
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function ShowDriveType(Drive)

	Dim S 	
	
	Select Case Drive.DriveType
	Case DriveTypeRemovable
		S = "Wechselmedium"
	Case DriveTypeFixed
		S = "Fest"
	Case DriveTypeNetwork
		S = "Netzwerk"
	Case DriveTypeCDROM
		S = "CD-ROM"
	Case DriveTypeRAMDisk
		S = "RAM-Laufwerk"
	Case Else
		S = "Unbekannt"
	End Select

	ShowDriveType = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' ShowFileAttr
'
' Zweck: 
'
' Erstellt eine Zeichenfolge, die Datei- oder Ordnerattribute
' beschreibt.
'
' Zeigt Folgendes 
'
' - File.Attributes
' - Folder.Attributes
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function ShowFileAttr(File) ' File kann Datei oder Ordner sein

	Dim S 	
	Dim Attr
	
	Attr = File.Attributes

	If Attr = 0 Then
		ShowFileAttr = "Normal"
		Exit Function
	End If

	If Attr And FileAttrDirectory  Then S = S & "Verzeichnis "
	If Attr And FileAttrReadOnly   Then S = S & "Schreibgeschützt"
	If Attr And FileAttrHidden     Then S = S & "Versteckt"
	If Attr And FileAttrSystem     Then S = S & "System"
	If Attr And FileAttrVolume     Then S = S & "Datenträger"
	If Attr And FileAttrArchive    Then S = S & "Archiv"
	If Attr And FileAttrAlias      Then S = S & "Alias"
	If Attr And FileAttrCompressed Then S = S & "Komprimiert"

	ShowFileAttr = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' GenerateDriveInformation
'
' Zweck: 
'
' Erstellt eine Zeichenfolge, die den aktuellen Status der verfügbaren
' Laufwerke beschreibt.
'
' Zeigt Folgendes 
'
' - FileSystemObject.Drives 
' - Iteration der Drives-Auflistung
' - Drives.Count
' - Drive.AvailableSpace
' - Drive.DriveLetter
' - Drive.DriveType
' - Drive.FileSystem
' - Drive.FreeSpace
' - Drive.IsReady
' - Drive.Path
' - Drive.SerialNumber
' - Drive.ShareName
' - Drive.TotalSize
' - Drive.VolumeName
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateDriveInformation(FSO)

	Dim Drives
	Dim Drive
	Dim S

	Set Drives = FSO.Drives

	S = "Anzahl der Laufwerke:" & TabStop & Drives.Count & NewLine & NewLine

	' Erstellt die erste Zeile des Berichts.
	S = S & String(2, TabStop) & "Laufwerk" 
	S = S & String(3, TabStop) & "Datei" 
	S = S & TabStop & "Gesamter"
	S = S & TabStop & "Freier"
	S = S & TabStop & "Verfügbarer"
	S = S & TabStop & "Seriennummer" & NewLine

	' Erstellt die zweite Zeile des Berichts.
	S = S & "Laufwerkbuchstabe"
	S = S & TabStop & "Pfad"
	S = S & TabStop & "Typ"
	S = S & TabStop & "Bereit ?"
	S = S & TabStop & "Name"
	S = S & TabStop & "System"
	S = S & TabStop & "Speicherplatz"
	S = S & TabStop & "Speicherplatz"
	S = S & TabStop & "Speicherplatz"
	S = S & TabStop & "Nummer" & NewLine	

	' Trennlinie.
	S = S & String(105, "-") & NewLine

	For Each Drive In Drives

		S = S & Drive.DriveLetter
		S = S & TabStop & Drive.Path
		S = S & TabStop & ShowDriveType(Drive)
		S = S & TabStop & Drive.IsReady

		If Drive.IsReady Then 		
			If DriveTypeNetwork = Drive.DriveType Then
				S = S & TabStop & Drive.ShareName 
			Else
				S = S & TabStop & Drive.VolumeName 
			End If    

			S = S & TabStop & Drive.FileSystem
			S = S & TabStop & Drive.TotalSize
			S = S & TabStop & Drive.FreeSpace
			S = S & TabStop & Drive.AvailableSpace
			S = S & TabStop & Hex(Drive.SerialNumber)

		End If    

		S = S & NewLine

	Next  
	
	GenerateDriveInformation = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' GenerateFileInformation
'
' Zweck: 
'
' Erstellt eine Zeichenfolge, die den aktuellen Status einer Datei
' beschreibt.
'
' Zeigt Folgendes 
'
' - File.Path
' - File.Name
' - File.Type
' - File.DateCreated
' - File.DateLastAccessed
' - File.DateLastModified
' - File.Size
' 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateFileInformation(File)

	Dim S

	S = NewLine & "Pfad:" & TabStop & File.Path
	S = S & NewLine & "Name:" & TabStop & File.Name
	S = S & NewLine & "Typ:" & TabStop & File.Type
	S = S & NewLine & "Attribute:" & TabStop & ShowFileAttr(File)
	S = S & NewLine & "Erstellt:" & TabStop & File.DateCreated
	S = S & NewLine & "Letzter Zugriff:" & TabStop & File.DateLastAccessed
	S = S & NewLine & "Letzte Änderung:" & TabStop & File.DateLastModified
	S = S & NewLine & "Größe" & TabStop & File.Size & NewLine

	GenerateFileInformation = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' GenerateFolderInformation
'
' Zweck:
'
' Erstellt eine Zeichenfolge, die den aktuellen Status eines Ordners
' beschreibt.
'
' Zeigt Folgendes
'
' - Folder.Path
' - Folder.Name
' - Folder.DateCreated
' - Folder.DateLastAccessed
' - Folder.DateLastModified
' - Folder.Size
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateFolderInformation(Folder)

	Dim S

	S = "Pfad:" & TabStop & Folder.Path
	S = S & NewLine & "Name:" & TabStop & Folder.Name
	S = S & NewLine & "Attribute:" & TabStop & ShowFileAttr(Folder)
	S = S & NewLine & "Erstellt:" & TabStop & Folder.DateCreated
	S = S & NewLine & "Letzter Zugriff:" & TabStop & Folder.DateLastAccessed
	S = S & NewLine & "Letzte Änderung:" & TabStop & Folder.DateLastModified
	S = S & NewLine & "Größe:" & TabStop & Folder.Size & NewLine

	GenerateFolderInformation = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' GenerateAllFolderInformation
'
' Zweck: 
'
' Erstellt eine Zeichenfolge, die den aktuellen Status eines Ordners mit
' allen Dateien und Unterordnern beschreibt.
'
' Zeigt Folgendes 
'
' - Folder.Path
' - Folder.SubFolders
' - Folders.Count
' 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateAllFolderInformation(Folder)

	Dim S
	Dim SubFolders
	Dim SubFolder
	Dim Files
	Dim File

	S = "Ordner:" & TabStop & Folder.Path & NewLine & NewLine

	Set Files = Folder.Files

	If 1 = Files.Count Then
		S = S & "Es ist 1 Datei vorhanden" & NewLine
	Else
		S = S & "Es sind " & Files.Count & "Dateien vorhanden" & NewLine
	End If

	If Files.Count <> 0 Then

		For Each File In Files
			S = S & GenerateFileInformation(File)
		Next

	End If

	Set SubFolders = Folder.SubFolders

	If 1 = SubFolders.Count Then
		S = S & NewLine & "Es ist 1 Unterordner vorhanden" & NewLine & NewLine
	Else
		S = S & NewLine & "Es sind" & SubFolders.Count & "Unterordner vorhanden" & NewLine & NewLine
	End If

	If SubFolders.Count <> 0 Then

		For Each SubFolder In SubFolders
			S = S & GenerateFolderInformation(SubFolder)
		Next

		S = S & NewLine

		For Each SubFolder In SubFolders
			S = S & GenerateAllFolderInformation(SubFolder)
		Next

	End If

	GenerateAllFolderInformation = S

End Function

Function CopyAllFolderFiles(Folder, FSO, bForceRepl, bSkip, stSkippedItem)
'FileSystemObject.CopyFile "c:\EigeneDateien\Briefe\*.doc", "c:\temp\"
	Dim S
	Dim SubFolders
	Dim SubFolder
	Dim Files
	Dim File
    'Dim Name
    Dim NewName
    Dim NewFile
    Dim Proceed

    Proceed = false
'WScript.Echo Folder.Path '@debug

	S = "Ordner:" & TabStop & Folder.Path '& NewLine & NewLine
        'S= ""

	Set Files = Folder.Files

	If 1 = Files.Count Then
		'S = S & "Es ist 1 Datei vorhanden" & NewLine
	Else
		'S = S & "Es sind " & Files.Count & "Dateien vorhanden" & NewLine
	End If

	If Files.Count <> 0 Then

		For Each File In Files
			'S = S & GenerateFileInformation(File)
                        'Name = File.Path
                        'NewName = stTargetPath & Mid(Name, Len(stFilePath)+1)
                        NewName = ReplacePath(File)
'WScript.Echo File.Path & " -> " & NewName
                        If FSO.FileExists(NewName) Then
                                Set NewFile = FSO.getFile(NewName)
                                If NewFile.Attributes And FileAttrReadOnly Then
                                ' @todo: Add functions / par for replace ReadOnly files !
                                '        For now we DO replace them anyway
                                        NewFile.Attributes = NewFile.Attributes - FileAttrReadOnly
                                        Proceed = true
                                End If

                                ' @todo: Add functions / par for replace NEWER files only !
                                '        For now we DO replace them if newer
                                If NewFile.DateLastModified >= File.DateLastModified Then
                                        Proceed = false
                                Else
                                        Proceed = true
                                End If
                        Else
                                Proceed = true
                        End If

                        If Proceed Then
'WScript.Echo File.Path & " -> " & NewName
                                If bForceRepl Then
									If FSO.FileExists(NewName) Then FSO.DeleteFile NewName, true
								End If
								File.Copy NewName, true
                        End If

                        NewName = ""
                        Proceed = false
		Next

	End If

	Set SubFolders = Folder.SubFolders

	If 1 = SubFolders.Count Then
		'S = S & NewLine & "Es ist 1 Subdirectory found" & NewLine & NewLine
	Else
		'S = S & NewLine & "Es sind" & SubFolders.Count & "Subdirectories found" & NewLine & NewLine
	End If

	If SubFolders.Count <> 0 Then
		'S = S & NewLine

		For Each SubFolder In SubFolders
			If (bSkip And SubFolder.Name = stSkippedItem) Then
				' Do nothing
			Else
				'S = S & GenerateFolderInformation(SubFolder)
				'CopyAllFolderFiles SubFolder, FSO, bForceRepl, false, ""

				S = S & NewLine
				S = S & "Ordner:" & TabStop & SubFolder.Path
				
				NewName = ReplaceFolderPath(SubFolder)
				If FSO.FolderExists(NewName) Then
				'S = S & GenerateAllFolderInformation(SubFolder)
'WScript.Echo SubFolder.Path '@debug
					If bForceRepl Then
						FSO.DeleteFolder NewName, true
						'CopyAllFolderFiles SubFolder, FSO, bForceRepl, false, ""
						SubFolder.Copy(NewName)
					Else
						CopyAllFolderFiles SubFolder, FSO, false, false, ""
					End If 
				Else
'WScript.Echo SubFolder.Path '@debug
						SubFolder.Copy(NewName)
				End If
			End If
		Next

	End If

	CopyAllFolderFiles = S
End Function

Private Function ReplaceFilePath(File)
        Dim Name
        Dim NewName

        Name = File.Path
        NewName = stTargetPath & Mid(Name, Len(stFilePath)+1)
        ReplaceFilePath = NewName
End Function

Private Function ReplacePath(File)
' @todo: For JScript: Use ReplacePath(File) and ReplacePath(Folder) (Overwriting !!)
        ReplacePath = ReplaceFilePath(File)
End Function

Private Function ReplaceFolderPath(Folder)
        Dim Name
        Dim NewName

        Name = Folder.Path
        NewName = stTargetPath & Mid(Name, Len(stFilePath)+1)
        ReplaceFolderPath = NewName
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' GenerateCopyInformation
'
' Zweck: 
'
' Erstellt eine Zeichenfolge, die den aktuellen Status des Ziel-Ordners
' mit allen Dateien und Unterordnern beschreibt.
'
' Zeigt Folgendes 
'
' - FileSystemObject.DriveExists
' - FileSystemObject.FolderExists
' - FileSystemObject.GetFolder
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateCopyInformation(FSO, WshShell, bForceReplace, bSkip, stSkippedItem)
'On Error Resume Next
	Dim SourceFolder
    Dim TargetFolder
	Dim s

	If stDrive <> "\" Then
		If Not FSO.DriveExists(stDrive) Then 
			s = "Source drive (" & stDrive & ") is not ready."
			'WScript.Echo s
			WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconStop
			Exit Function
		End If
	End If
	If stTargetDrive <> "\" Then
		If Not FSO.DriveExists(stTargetDrive) Then			
			s = "Target drive (" & stTargetDrive & ") is not ready."
			'WScript.Echo s
			WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconStop
			Exit Function
		End If
	End If
	If Not FSO.FolderExists(stFilePath) Then
		s = "Source Drive does not exist."
		'WScript.Echo s
		WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconExclamation
		Exit Function
	End If
	
	Set SourceFolder = FSO.GetFolder(stFilePath)
        If Not FSO.FolderExists(stTargetPath) Then
                Set TargetFolder = FSO.CreateFolder(stTargetPath)
        Else
                Set TargetFolder = FSO.GetFolder(stTargetPath)
        End If
    If Err.Number <> 0 Then
		WScript.Echo "Fehler " & Err.Number & ": " & Err.Description
		Exit Function		
		'Err.Clear
    End If
    
	'GenerateCopyInformation = GenerateAllFolderInformation(SourceFolder)
        GenerateCopyInformation = CopyAllFolderFiles(SourceFolder, FSO, bForceReplace, bSkip, stSkippedItem)
'WScript.Echo SourceFolder.Path & " -> " & stTargetPath & "\"
        'FSO.CopyFolder SourceFolder.Path, stTargetPath & "\"
	Exit Function

'GenerateTestInformation_Err:	
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' DeleteTestDirectory
'
' Purpose:
'
' Cleans the Source folder.
'
' Shows results
'
' - FileSystemObject.GetFolder
' - FileSystemObject.DeleteFile
' - FileSystemObject.DeleteFolder
' - Folder.Delete
' - File.Delete
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub DeleteTestDirectory(FSO)
	Dim SourceFolder
	Dim SubFolder
	Dim File

	' Zwei Möglichkeiten, eine Datei zu löschen:

	FSO.DeleteFile(stFilePath & "\Phish\BathtubGin.txt")

	Set File = FSO.GetFile(stFilePath & "\Phish\LawnBoy.txt")
	File.Delete



	' Zwei Möglichkeiten, einen Ordner zu löschen:

	FSO.DeleteFolder(stFilePath & "\Phish")

	FSO.DeleteFile(stFilePath & "\Liesmich.txt")

	Set SourceFolder = FSO.GetFolder(stFilePath)
	SourceFolder.Delete
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' BuildTestDirectory
'
' @Deprecated: unused
' Purpose:
'
' Erstellt eine Verzeichnishierarchie, um FileSystemObject zu
' beschreiben.
'
' Die Hierarchie wird in dieser Reihenfolge erstellt:
'
' C:\Test
' C:\Test\Liesmich.txt
' C:\Test\Phish
' C:\Test\Phish\BathtubGin.txt
' C:\Test\Phish\LawnBoy.txt
'
'
' Zeigt Folgendes 
'
' - FileSystemObject.DriveExists
' - FileSystemObject.FolderExists
' - FileSystemObject.CreateFolder
' - FileSystemObject.CreateTextFile
' - Folders.Add
' - Folder.CreateTextFile
' - TextStream.writeLine
' - TextStream.Close
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function BuildTestDirectory(FSO)
	Dim SourceFolder
	Dim SubFolders
	Dim SubFolder
	Dim TextStream

	' Bricht ab, wenn (a) das Laufwerk nicht vorhanden oder (b) das zu erstellende Verzeichnis bereits vorhanden ist.

	If Not FSO.DriveExists(stDrive) Then
		BuildTestDirectory = false
		Exit Function
	End If

	If FSO.FolderExists(stFilePath) Then
		BuildTestDirectory = false
		Exit Function
	End If

	Set SourceFolder = FSO.CreateFolder(stFilePath)

	Set TextStream = FSO.CreateTextFile(stFilePath & "\Readme.txt")
	TextStream.writeLine("My Music")
	TextStream.Close

	Set SubFolders = SourceFolder.SubFolders

	Set SubFolder = SubFolders.Add("Phish")

	'CreateLyrics SubFolder	

	BuildTestDirectory = true
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' The main routine
'
' Zunächst wird ein Zielverzeichnis mit einigen Unterordnern und
' Dateien erstellt.
' Anschließend werden Informationen über die verfügbaren
' Festplattenlaufwerke sowie das Testverzeichnis erstellt.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Main ()
	Dim WshShell
	Dim FSO
	
	Dim objArgs
	Dim s
	Dim i
	Dim iShow
	Dim bKeep
	Dim bForceReplace
	Dim bSkip
	Dim stSkippedItem
	
	iShow = ShowResultNone
	bKeep = false
	bForceReplace = false
	bSkip = false
	
	' Set Pseudo-Constants (values that cannot be declared const for some reason)
	TabStop = Chr(9)
	NewLine = Chr(10)
	
	Set WshShell = Wscript.CreateObject("Wscript.Shell")
	Set FSO = CreateObject("Scripting.FileSystemObject")

	Set objArgs = Wscript.Arguments
	For i = 0 to objArgs.Count - 1
          'Wscript.Echo objArgs(i)
          If i=0 Then 
          	If objArgs(i) = "-info" Then
				iShow = ShowResultInfo
				bKeep = true
			Else
				s = objArgs(i)
				stFilePath = objArgs(i)
			End If
		  ElseIf i=1 Then
		  	If objArgs(i) = "-info" Then
				iShow = ShowResultInfo
				bKeep = true
			Else
				s = s & NewLine & objArgs(i)
				stTargetPath = objArgs(i)
			End If
          Else
			If objArgs(i) = "-test" Then 
			 	iShow = ShowResultAll
			ElseIf objArgs(i) = "-paths" Then 
				iShow = ShowResultStart
			ElseIf objArgs(i) = "-result" Then
				iShow = ShowResultFinal
			ElseIf objArgs(i) = "-keepresult" Then
				iShow = ShowResultFinal
				bKeep = true
			ElseIf objArgs(i) = "-forcereplace" Then 
' Forcing an explicit REPLACE of existing Folders (for Backup on certain media like DVD-RAM)
				bForceReplace = true
			ElseIf objArgs(i) = "-skip" Then
				bSkip = true
			ElseIf objArgs(i) = "-skipFolder" Then
				bSkip = true
			Else
				If (bSkip) Then
					If stSkippedItem = "" Then
						stSkippedItem = objArgs(i)
					Else
						s = s & NewLine & objArgs(i)
					End If
				Else
					s = s & NewLine & objArgs(i)
				End If
			End If
		  End If			
	Next
		
	If iShow = ShowResultInfo Then
			s="Version " + AppVersion
			
			If bKeep Then
				WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconInfo
			Else
				s = s & NewLine & NewLine & "(Finished in " & DelayDuration & " seconds)"
				WshShell.Popup s, 15, AppTitle, PopUpButtonOK + PopUpIconInfo
			End If		
		Exit Sub
	Else
		If iShow = ShowResultAll Or iShow = ShowResultStart Then 
			'WScript.Echo s
			s = s & NewLine & NewLine & "(Continues in " & DelayDuration & " seconds)"
			'if bKeep Then
			'	WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconInfo
			'Else
				WshShell.Popup s, DelayDuration, AppTitle, PopUpButtonOK + PopUpIconInfo
			'End If
		End If
	End If
	s = ""
	
	stDrive = Left(stFilePath, 1)
	stTargetDrive = Left(stTargetPath, 1)
	
'Wscript.Echo pars

	'If Not BuildTestDirectory(FSO) Then
	'	Wscript.Echo "Test folder cannot be created or already exists. Cannot continue."
	'	Exit Sub
'    Else
	'	Wscript.Echo "Test folder created succesfully."
	'End If

	'Wscript.Echo GenerateDriveInformation(FSO) & NewLine & NewLine

	s=GenerateCopyInformation(FSO, WshShell, bForceReplace, bSkip, stSkippedItem) ' & NewLine & NewLine
	If iShow = ShowResultAll Or iShow = ShowResultFinal Then 
		If Len(s)>0 Then 					
			s = "Copy succeeded." & NewLine & s
			If (bSkip And stSkippedItem <> "") Then
				s = s & NewLine & "Skipped: " & stSkippedItem
			End If
			'WScript.Echo s
			If bKeep Then
				WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconInfo
			Else
				s = s & NewLine & NewLine & "(Ende nach " & MessageDisplayDuration & " Sekunden)"
				WshShell.Popup s, MessageDisplayDuration, AppTitle, PopUpButtonOK + PopUpIconInfo
			End If
		End If
	End If

	'Print GetLyrics(FSO) & NewLine & NewLine

	'DeleteTestDirectory(FSO)
End Sub

'--------------------------------------------------------------------------------
