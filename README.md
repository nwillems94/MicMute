# MicMute
Simple utility to mute system microphone using C# and Powershell.  

I have configured this to run in place of the "My Computer" shortcut and remapped my keyboard accordingly. 
Something like the following regedit overwrites the shortcut:  

`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\17]
"ShellExecute"="wscript \"C:\\Users\\path\\to\\scripts\\MuteMic.js\""`
