on error resume next
Set objShell = CreateObject("Wscript.Shell")
Call objShell.Run("Powershell.exe" + " -ExecutionPolicy Bypass" + " C:\ProgramData\PlexusStartup\LocalAdminCheck.ps1" + " -windowstyle hidden", Hide_Window, true)
Call objShell.Run("Powershell.exe" + " $script:balloon.Dispose()" + " -windowstyle hidden", Hide_Window, true)

the above code need to save in text file and name it RunAdminCheck.vbs
