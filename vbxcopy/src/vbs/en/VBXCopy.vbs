'--------------------------------------------------------------------------------
' $Id: VBXCopy.vbs,v 1.9 2010/02/07 23:02:41 keilw Exp $
'--------------------------------------------------------------------------------

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' VBXCopy
'
' Version 1.0.11
'
' Copyright 1999-2010 Creative Arts & Technologies. All Rights reserved.
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
Const AppVersion = "1.0.11"
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
		S = "Removable Media"
	Case DriveTypeFixed
		S = "Fixed"
	Case DriveTypeNetwork
		S = "Network"
	Case DriveTypeCDROM
		S = "CD-ROM"
	Case DriveTypeRAMDisk
		S = "RAM-Drive"
	Case Else
		S = "Unknown"
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

Function ShowFileAttr(File) ' File can be FileFolder or Folder

	Dim S 	
	Dim Attr
	
	Attr = File.Attributes

	If Attr = 0 Then
		ShowFileAttr = "Normal"
		Exit Function
	End If

	If Attr And FileAttrDirectory  Then S = S & "Folder "
	If Attr And FileAttrReadOnly   Then S = S & "Readonly"
	If Attr And FileAttrHidden     Then S = S & "Hidden"
	If Attr And FileAttrSystem     Then S = S & "System"
	If Attr And FileAttrVolume     Then S = S & "Drive"
	If Attr And FileAttrArchive    Then S = S & "Archive"
	If Attr And FileAttrAlias      Then S = S & "Alias"
	If Attr And FileAttrCompressed Then S = S & "Compressed"

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
	S = S & String(2, TabStop) & "Drive" 
	S = S & String(3, TabStop) & "File" 
	S = S & TabStop & "Gesamter"
	S = S & TabStop & "Freier"
	S = S & TabStop & "Verfügbarer"
	S = S & TabStop & "Seriennummer" & NewLine

	' Erstellt die zweite Zeile des Berichts.
	S = S & "Drive Letter"
	S = S & TabStop & "Path"
	S = S & TabStop & "Type"
	S = S & TabStop & "Ready ?"
	S = S & TabStop & "Name"
	S = S & TabStop & "System"
	S = S & TabStop & "Memory"
	S = S & TabStop & "Memory"
	S = S & TabStop & "Memory"
	S = S & TabStop & "Number" & NewLine	

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
	S = NewLine & "Path:" & TabStop & File.Path
	S = S & NewLine & "Name:" & TabStop & File.Name
	S = S & NewLine & "Type:" & TabStop & File.Type
	S = S & NewLine & "Attribute:" & TabStop & ShowFileAttr(File)
	S = S & NewLine & "Created:" & TabStop & File.DateCreated
	S = S & NewLine & "Last Accessed:" & TabStop & File.DateLastAccessed
	S = S & NewLine & "Last Modified:" & TabStop & File.DateLastModified
	S = S & NewLine & "Size" & TabStop & File.Size & NewLine

	GenerateFileInformation = S

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' GenerateFolderInformation
'
' Purpose:
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
' @todo localize @loc
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

	S = "Folder:" & TabStop & Folder.Path & NewLine & NewLine

	Set Files = Folder.Files

	If 1 = Files.Count Then
		S = S & "Found 1 File" & NewLine
	Else
		S = S & "" & Files.Count & " Files found" & NewLine
	End If

	If Files.Count <> 0 Then

		For Each File In Files
			S = S & GenerateFileInformation(File)
		Next

	End If

	Set SubFolders = Folder.SubFolders

	If 1 = SubFolders.Count Then
		S = S & NewLine & "1 Subdirectory found" & NewLine & NewLine
	Else
		S = S & NewLine & "" & SubFolders.Count & " Subdirectories found" & NewLine & NewLine
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CopyAllFolderFiles
'
' Zweck: 
'
' Kopiert den Ordner und alle darin enthaltenen Dateien und Unterordnern.
' (Rekursiv)
' 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function CopyAllFolderFiles(Folder, FSO, bForceRepl, bSkipOld, bSkipBAK, bSkipCVS)
'FileSystemObject.CopyFile "c:\My Documents\Letters\*.doc", "c:\temp\"

	Dim S
	Dim SubFolders
	Dim SubFolder
	Dim Files
	Dim File
        'Dim Name
        Dim NewName
        Dim NewFile
        Dim Proceed
	Dim ProceedAll

        ProceedAll = False
        Proceed = False
'WScript.Echo Folder.Path '@debug

	If bSkipOld Then
' TODO @todo add Skips for BAK and CVS here, too (later)		
		If Folder.Name = "Old" Then
			' do Nothing (ProceedAll remains false)
		Else
			ProceedAll = True
		End If
	Else
		ProceedAll = True
	End If

	If ProceedAll Then
		S = "Folder:" & TabStop & Folder.Path '& NewLine & NewLine
        'S= ""

	Set Files = Folder.Files

	If 1 = Files.Count Then
			If gbDebug Then
				S = S & NewLine & "1 File found" & NewLine
			End If
	Else
			If gbDebug Then
				S = S & NewLine & "" & Files.Count & " Files found" & NewLine
			End If
	End If

	If Files.Count <> 0 Then

		For Each File In Files
			'S = S & GenerateFileInformation(File)
                        'Name = File.Path
                        'NewName = TestTargetPath & Mid(Name, Len(TestFilePath)+1)
                        NewName = ReplacePath(File)
'WScript.Echo File.Path & " -> " & NewName
                        If FSO.FileExists(NewName) Then
                                Set NewFile = FSO.getFile(NewName)
                                If NewFile.Attributes And FileAttrReadOnly Then
                                ' @todo: Add functions / par for replace ReadOnly files !
                                '        For now we DO replace them anyway
                                        NewFile.Attributes = NewFile.Attributes - FileAttrReadOnly
                                        Proceed = True
                                End If

                                ' @todo: Add functions / par for replace NEWER files only !
                                '        For now we DO replace them if newer
                                If NewFile.DateLastModified >= File.DateLastModified Then
                                        Proceed = False
                                Else
                                        Proceed = True
                                End If
                        Else
                                Proceed = True
                        End If

                        If Proceed Then
'WScript.Echo File.Path & " -> " & NewName
                                If bForceRepl Then
									If FSO.FileExists(NewName) Then FSO.DeleteFile NewName, true
									File.Copy NewName, true
                                Else
									File.Copy NewName, true
									'File.Copy TestTargetPath
								End If
                        End If

                        NewName = ""
                        Proceed = False
		Next

	End If

	Set SubFolders = Folder.SubFolders

	If 1 = SubFolders.Count Then
			If gbDebug Then
				S = S & NewLine & "1 Subdirectory found" & NewLine
			End If
	Else
			If gbDebug Then
				S = S & NewLine & "" & SubFolders.Count & " Subdirectories found" & NewLine
			End If
	End If

	If SubFolders.Count <> 0 Then

		For Each SubFolder In SubFolders
			'S = S & GenerateFolderInformation(SubFolder)
	                        'CopyAllFolderFiles SubFolder, FSO, bForceRepl, bSkipOld, bSkipBAK, bSkipCVS
                        S = S & NewLine
                        S = S & "Folder:" & TabStop & SubFolder.Path
		Next

		'S = S & NewLine

		For Each SubFolder In SubFolders
                        NewName = ReplaceFolderPath(SubFolder)
                        If FSO.FolderExists(NewName) Then
			        'S = S & GenerateAllFolderInformation(SubFolder)
'WScript.Echo SubFolder.Path '@debug
							If bForceRepl Then
	'WScript.Echo "Repl for " & SubFolder.Path '@debug
                                FSO.DeleteFolder NewName, true
	                                'CopyAllFolderFiles SubFolder, FSO, bForceRepl, bSkipOld, bSkipBAK, bSkipCVS
                                SubFolder.Copy(NewName)
                            Else
	'WScript.Echo "NO Repl for " & SubFolder.Path '@debug
					CopyAllFolderFiles SubFolder, FSO, false, bSkipOld, bSkipBAK, bSkipCVS
                            End If 
                        Else
'WScript.Echo SubFolder.Path '@debug
	'WScript.Echo SubFolder.Name '@debug
	'WScript.Echo NewName '@debug
				    If bSkipOld Then
				    	If Not (SubFolder.Name = "OLD" Or SubFolder.Name = "Old" Or SubFolder.Name = "old" Or SubFolder.Name = "_old") Then
				    		If bSkipBAK Then
				    			If Not (SubFolder.Name = "BAK" Or SubFolder.Name = "Bak" Or SubFolder.Name = "bak") Then
				    				SubFolder.Copy(NewName)
				    			End If	
				    		Else
				    			If bSkipCVS Then
				    				If Not (SubFolder.Name = "CVS") Then 'here only UPPERCASE (used by CVS)
                                SubFolder.Copy(NewName)
                        End If
				    			Else
				    				SubFolder.Copy(NewName)
				    			End If
				    		End If			    	
				    	End If
				    Else
	                            	SubFolder.Copy(NewName)
	                            End If	                            
	                        End If
		Next

	End If
	Else
		S = ""		
	End If

	CopyAllFolderFiles = S
End Function

Private Function ReplaceFilePath(File)
        Dim Name
        Dim NewName

        Name = File.Path
        NewName = TestTargetPath & Mid(Name, Len(TestFilePath)+1)
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
        NewName = TestTargetPath & Mid(Name, Len(TestFilePath)+1)
        ReplaceFolderPath = NewName
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' GenerateResult
'
' Zweck: 
'
' Erstellt eine Zeichenfolge, die den aktuellen Status des Ordners
' C:\Test mit allen Dateien und Unterordnern beschreibt.
'
' Zeigt Folgendes 
'
' - FileSystemObject.DriveExists
' - FileSystemObject.FolderExists
' - FileSystemObject.GetFolder
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateResult(FSO, WshShell, bForceReplace, bSkipOld, bSkipBAK, bSkipCVS)
'On Error Resume Next
	Dim TestFolder
    Dim TargetFolder
	Dim s

	If TestDrive <> "\" Then
		If Not FSO.DriveExists(TestDrive) Then 
			s = "Source Drive (" & TestDrive & ") is not ready."
			'WScript.Echo s
			WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconStop
			Exit Function
		End If
	End If
	If TestTargetDrive <> "\" Then
		If Not FSO.DriveExists(TestTargetDrive) Then			
			s = "Target Drive (" & TestTargetDrive & ") is not ready."
			'WScript.Echo s
			WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconStop
			Exit Function
		End If
	End If
	If Not FSO.FolderExists(TestFilePath) Then
		s = "Source Drive does not exist."
		'WScript.Echo s
		WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconExclamation
		Exit Function
	End If
	
	Set TestFolder = FSO.GetFolder(TestFilePath)
        If Not FSO.FolderExists(TestTargetPath) Then
                Set TargetFolder = FSO.CreateFolder(TestTargetPath)
        Else
                Set TargetFolder = FSO.GetFolder(TestTargetPath)
        End If
    If Err.Number <> 0 Then
		WScript.Echo "Error " & Err.Number & ": " & Err.Description
		Exit Function		
		'Err.Clear
    End If
    
	'GenerateResult = GenerateAllFolderInformation(TestFolder)
        GenerateResult = CopyAllFolderFiles(TestFolder, FSO, bForceReplace, bSkipOld, bSkipBAK, bSkipCVS)
'WScript.Echo TestFolder.Path & " -> " & TestTargetPath & "\"
        'FSO.CopyFolder TestFolder.Path, TestTargetPath & "\"
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

	Dim TestFolder
	Dim SubFolder
	Dim File

	' Zwei Möglichkeiten, eine Datei zu löschen:

	FSO.DeleteFile(TestFilePath & "\Phish\BathtubGin.txt")

	Set File = FSO.GetFile(TestFilePath & "\Phish\LawnBoy.txt")
	File.Delete



	' Zwei Möglichkeiten, einen Ordner zu löschen:

	FSO.DeleteFolder(TestFilePath & "\Phish")

	FSO.DeleteFile(TestFilePath & "\Readme.txt")

	Set TestFolder = FSO.GetFolder(TestFilePath)
	TestFolder.Delete

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' BuildTestDirectory
'
' Zweck:
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
	
	Dim TestFolder
	Dim SubFolders
	Dim SubFolder
	Dim TextStream

	' Bricht ab, wenn (a) das Laufwerk nicht vorhanden oder (b) das zu erstellende Verzeichnis bereits vorhanden ist.

	If Not FSO.DriveExists(TestDrive) Then
		BuildTestDirectory = False
		Exit Function
	End If

	If FSO.FolderExists(TestFilePath) Then
		BuildTestDirectory = False
		Exit Function
	End If

	Set TestFolder = FSO.CreateFolder(TestFilePath)

	Set TextStream = FSO.CreateTextFile(TestFilePath & "\Readme.txt")
	TextStream.writeLine("My Music")
	TextStream.Close

	Set SubFolders = TestFolder.SubFolders

	Set SubFolder = SubFolders.Add("Phish")

	'CreateLyrics SubFolder	

	BuildTestDirectory = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Die Hauptroutine
'
' Zunächst wird ein Testverzeichnis mit einigen Unterordnern und
' Dateien erstellt.
' Anschließend werden Informationen über die verfügbaren
' Festplattenlaufwerke sowie das Testverzeichnis erstellt und danach
' alles wieder entfernt.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Main ()
	Dim WshShell
	Dim FSO
	
	Dim objArgs
	Dim s
	Dim I
	Dim iShow
	Dim bKeep
	Dim bForceReplace
	Dim bSkipOld 'TODO @todo use Array (skipList) if that works to dim in VBScript?
	Dim bSkipBAK
	Dim bSkipCVS
	
	' Einrichten globaler Daten.
	TabStop = Chr(9)
	NewLine = Chr(10)
	gbDebug = false
	
	iShow = ShowResultNone
	bKeep = false
	bForceReplace = false
	
	bSkipOld = false
	bSkipBAK = false
	bSkipCVS = false
	
	Set WshShell = Wscript.CreateObject("Wscript.Shell")
	Set FSO = CreateObject("Scripting.FileSystemObject")

	Set objArgs = Wscript.Arguments
	For I = 0 to objArgs.Count - 1
          'Wscript.Echo objArgs(I)
          If I=0 Then 
			If objArgs(I) = "-info" Then
				iShow = ShowResultInfo
				bKeep = true
			Else
				s = objArgs(I)
				TestFilePath = objArgs(I)
			End If
		  ElseIf I=1 Then
			If objArgs(I) = "-info" Then
				iShow = ShowResultInfo
				bKeep = true
			Else			
				s = s & NewLine & objArgs(I)
				TestTargetPath = objArgs(I)
			End If
          Else
			If objArgs(I) = "-test" Then 
			 	iShow = ShowResultAll
			ElseIf objArgs(I) = "-paths" Then 
				iShow = ShowResultStart
			ElseIf objArgs(I) = "-result" Then
				iShow = ShowResultFinal
			ElseIf objArgs(I) = "-keepresult" Then 'TODO use Uppercase R? (or both for existing links!)
				iShow = ShowResultFinal
				bKeep = true
			ElseIf objArgs(I) = "-forcereplace" Then 'TODO use Uppercase R? (or both for existing links!)
' Forcing an explicit REPLACE of existing Folders (for Backup on certain media like DVD-RAM)
				bForceReplace = true
			ElseIf objArgs(I) = "-skipold" Then 'TODO use Uppercase O?
				bSkipOld = true
			ElseIf objArgs(I) = "-skipbak" Then 'TODO use Uppercase B or all BAK?
				bSkipBAK = true
			ElseIf objArgs(I) = "-skipcvs" Then 'TODO use Uppercase C or all CVS?
				bSkipCVS = true
			ElseIf objArgs(I) = "-debug" Then
				gbDebug = true
			Else
				s = s + NewLine + objArgs(I)
			End If
		  End If			
	Next
		
	If iShow = ShowResultInfo Then
			s="Version " + AppVersion
			
			If bKeep Then
				WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconInfo
			Else
				s = s & NewLine & NewLine & "(Finished in 15 Seconds)"
				WshShell.Popup s, 15, AppTitle, PopUpButtonOK + PopUpIconInfo
			End If
		
		Exit Sub
	Else
		If iShow = ShowResultAll Or iShow = ShowResultStart Then 
			'WScript.Echo s
			s = s & NewLine & NewLine & "(Continues in 5 Seconds)"
			'if bKeep Then
			'	WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconInfo
			'Else
				WshShell.Popup s, 5, AppTitle, PopUpButtonOK + PopUpIconInfo
			'End If
		End If
		
		s = ""
		
		TestDrive = Left(TestFilePath,1)
		TestTargetDrive = Left(TestTargetPath,1)
		
	'Wscript.Echo pars
	
		'If Not BuildTestDirectory(FSO) Then
		'	Wscript.Echo "Test folder cannot be created or may already exist. Unable to proceed."
		'	Exit Sub
	'    Else
		'	Wscript.Echo "Test folder successfully created."
		'End If
	
		'Wscript.Echo GenerateDriveInformation(FSO) & NewLine & NewLine
	
	s = GenerateResult(FSO, WshShell, bForceReplace, bSkipOld, bSkipBAK, bSkipCVS) ' & NewLine & NewLine
		If iShow = ShowResultAll Or iShow = ShowResultFinal Then 
			If Len(s)>0 Then 					
				s = "Copy succeeded." & NewLine & s
				'WScript.Echo s
				if bKeep Then
					WshShell.Popup s, , AppTitle, PopUpButtonOK + PopUpIconInfo
				Else
					s = s & NewLine & NewLine & "(Finished in 15 Seconds)"
					WshShell.Popup s, 15, AppTitle, PopUpButtonOK + PopUpIconInfo
				End If
			End If
		End If
	
		'Print GetLyrics(FSO) & NewLine & NewLine
	
		'DeleteTestDirectory(FSO)
	End If
End Sub


'--------------------------------------------------------------------------------
