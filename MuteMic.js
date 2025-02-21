var wshShell = new ActiveXObject("WScript.Shell");

// C# source code defining necessary methods and imports
var scriptloc = WScript.ScriptFullName.replace(".js",".cs")

// PowerShell commands
var psCommand = 
    // Import C# methods
    "$csCode = Get-Content -Path \'" + scriptloc + "\' -Raw; " +
    "Add-Type -TypeDefinition $csCode -ReferencedAssemblies \'System.Windows.Forms.dll\', \'System.Drawing.dll\'; " +
    // Execute the ToggleMuteMicrophone method and get last mute status
    "$wasMuted = [Program]::ToggleMuteMicrophone(); " +
    // Show the correct splashscreen for 1 second
    "$splashScreen = New-Object Program+SplashScreen -ArgumentList $wasMuted; " +
    "$splashScreen.Show(); " +
    "Start-Sleep -Seconds 1; " +
    "$splashScreen.Close();";

// Execute commmands without displaying a window
wshShell.Run('%SystemRoot%\\system32\\WindowsPowerShell\\v1.0\\powershell.exe -ExecutionPolicy Bypass -Command "' + psCommand + '"', 0, false);
