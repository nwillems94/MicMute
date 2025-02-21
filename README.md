# MicMute
Simple utility to mute system microphone using C# and Powershell and show an on-screen notification. Does **_not_** require installing any additional packages.  

You could create a shortcut to run the `.js` file and map a keyboard shortcut to it, but I have found that to be very slow to respond. 
Instead, I have configured this to run in place of the "My Computer" shortcut and remapped my keyboard accordingly (QMK). 
Something like the following regedit overwrites the "My Computer" shortcut:  

`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\17]
"ShellExecute"="wscript \"C:\\Users\\path\\to\\scripts\\MuteMic.js\""`
