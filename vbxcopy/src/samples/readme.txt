'$Id: readme.txt,v 1.2 2007/07/16 15:57:33 keilw Exp $

To use these samples you have to set the environment variable %BatchPath% for your (Windows)
system or user.

A common practice is, to define this in the "%CommonProgramFiles%" folder.
(e.g. "C:\Program Files\Common Files")

In a folder "Batch Files", "VBScript", "BAT" or "VBS".

If you then use the %CommonProgramFiles% you normally have to declare your %BatchPath% in the "User" section, as at least up to Windows XP the System variables have no access to it yet (it is declared as internal System variable)