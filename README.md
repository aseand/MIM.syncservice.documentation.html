# MIM.syncservice.portal.documentation.html
Script to generate attribute, flow and sync rules documentation for MIM/FIM to html file.
SynchronizationRules, MPR SET and workflow are included in html.

Require Lithnet.ResourceManagement.Client.dll and Microsoft.ResourceManagement.dll
Require user to have sync and service admin accesss (no SQL access i required)

Run MIM.syncservice.portal.documentation.html.ps1 on MIM server
Script will download Lithnet.ResourceManagement.Client nuget if dll is missing in current directory

![](preview.gif?raw=true "Title")
