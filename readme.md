The print-reset PowerShell script replaces the previous Microsoft FixIt that would perform a light or full print spooler reset as per
the following blog: https://blogs.technet.microsoft.com/askperf/2012/02/24/microsoft-fixit-for-printing/

MICROSOFT has deprecated all FixIts and replaced them with Troubleshooting Packs, therefore these are no longer available.

The PowerShell script detects the version of windows, deletes all user printer settings, and replaces it with known clean registry entries
for the print spooler and default print monitors.

The script must be run with either the -full or -light setting.

Keep in mind in certain situations that you will see what we have coined as "Ghost printers" where the print queue still exists even though
the printer connection itself is deleted and the drivers do not exist. We have identified two other registry locations that can be safely
cleaned that appears to completely wipe these ghosts. You must delete these as SYSTEM.

HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\DeviceClasses\{0ecef634-6ef0-472a-8085-5ad023ecbccd}
HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\SWD\PRINTENUM

Call this script from powershell directly as admin by running
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/DrDrrae/Powershell/master/Printers/print-reset.ps1') |iex
Reset-Printers -full -force
-OR-
Reset-Printers -light -force