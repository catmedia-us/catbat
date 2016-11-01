'$Id: readme.txt,v 1.4 2007/04/17 17:40:24 keilw Exp $

'VBXCopy is a tool by Creative Arts & Technologies
'which provides most features of the Windows/DOS command "XCopy" 
'plus useful enhancements for file synchronization and backup.
'
'Copyright © 1999-2016 Creative Arts & Technologies
'
'You can redistribute it and/or modify it under the terms of the
'Shared Source License as published by Creative Arts & Technologies
'Based on the Microsoft Shared Source License either version 1
'of the License or(at your option) any later version.
'
'You can use this Software for any non-commercial purpose,
'including distributing derivatives. Running your business operations
'would not be considered non-commercial.
'For commercial purposes, you can reference this software solely
'to assist in developing and testing your own software and hardware
'for this solution. You may not distribute this software in source
'or object form for commercial purposes under any circumstances.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Shared Source License for more details.
'
'You should have received a copy of the Shared Source License
'along with this program; if not, write to Creative Arts & Technologies
'at info@catmedia.us
'
'Creative Arts & Technologies Ltd. & Co. KG 
'Amalienstr. 71 
'D-80799 München 
'Germany
'
'General Partner: 
'Creative Arts & Technologies Ltd.
'Companies House Nr. 9724030 
'
'e-mail: info@catmedia.us
'WWW: http://www.catmedia.us/

@todo localize @loc - create readme_de.txt and others for additional languages
VBXCopy benutzen

Rufen Sie das Programm auf und wählen Sie ein Projekt. (Auf die Dateien
wird nur lesend zugegriffen - aber Sie können ja den Source lesen.)
Setzen Sie die Optionen:
    - "Alles in eine Datei":
        Wenn angekreuzt, wird eine große HTML-Datei ("HTML-Doc.htm")
            im Projektverzeichnis angelegt. 
            Gut für kleine Projekte.

        Wenn nicht angekreuzt, wird im Projektverzeichnis ein Ordner 
            "HTML-Doc" angelegt und für jede einzelne Source-Datei 
            eine eigene HTML-Datei geschrieben.

    - "Schrift":
        Das wird in die <FONT FACE="">-Angabe geschrieben. Wenn Sie
            eine andere Schrift mögen, geben Sie sie hier ein.

VBXCopy wertet ausschließlich Kommentare aus, die mit ' beginnen (Rem
also nicht); von denen auch nur die, die direkt am Anfang von Modulen
oder direkt hinter Prozedurdefinitionen stehen.
Folgende Keywords haben spezielle Bedeutungen:
'param
'return
'author
'version

Von allen Keywords außer "return" können mehrere in einer Prozedur
vorkommen; author und version können sowohl in Prozedurdefinitionen als
auch im Deklarationsbereich der Module verwendet werden.

Lange Beschreibungen können so geschrieben werden:
'param  Foo     Variable mit langer Beschreibung, die keinen
'               anderen Zweck hat als diese Methode zu erklären

Schauen Sie im Quellcode nach, das ist genauso dokumentiert.

Deinstallation
Löschen Sie VBXCopy.vbs, readme.txt und COPYING.txt und aus der Registry
den Zweig HKEY_CURRENT_USER\Software\CATMedia\VBXCopy.

Kontakt: Creative Arts & Technologies, info@catmedia.us
