# outlook-addon
A plug-in for Microsoft Outlook that allows reporting spam to a SpamExperts filtering cluster.

## Instructions for Building

Note that currently building the plug-in requires a very old version of Visual Studio, because of issues with DLL dependencies. This is something that needs to be fixed, so that current versions of Visual Studio can be used.

### Pre-requisites

 * Windows 7
 * Visual Studio 2010
 * Outlook 2010/2007
 * .Net Framework 4

### Steps

 1. Get a copy of the code (e.g. git clone)
 2. Find icon.ico and re-save with (e.g.) gimp to get the new icon version (new windows type version)
 3. Edit the "ReportSpamSetup\ReportSpamSetup.vdproj" file, and remove the following part:

    "SccProjectName" = "8:" 
    "SccLocalPath" = "8:"  
    "SccAuxPath" = "8:" 
    "SccProvider" = "8:"
 4. Save file
 5. Open Visual Studio 2010
 6. Load project from outlook/trunk
 7. Right click ReportSpam from Solution Explorer > Properties > Configuration properties > C/C++ > Add the following to PreProcessor Definitions `WINVER=0x0501` (or this to the Command line option `/D "_WIN32_WINNT=0x601` (depending on your windows version)
 8. Apply
 9. Click Build > configuration manager > tick all to be built
 10. Edit "Launch conditions" to set the version of .Net Framework to ".4" (right click "Detected Dependencies > click Framework)
 11. Build > Build Solution (make sure you set to release version and not debug)

The built files are located in `\ReportSpamSetup\Release\setup.exe` and `\ReportSpamSetup\Release64\setup.exe` (with the related .msi file).
