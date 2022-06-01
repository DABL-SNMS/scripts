#region 1. Definitie van variabelen
#region 1.1 Algemeen
$LW_Bron = "C:\Connect Folder\DABL Update"
$LW_Naam = "XRUpgrade"
$LW_Doel = "C:\Qastor projects"
$LW_Check = "C:\DABLSupport\Installed"
$LW_AntiVirus = "Symantec Endpoint Protection"
$LW_XRUpgraders = Import-Csv -Path $LW_Bron\$LW_Naam\$LW_Naam.csv -UseCulture
#endregion

#region 1.2 Regkeys
$LW_GPO_Reg_Pad = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WcmSvc\GroupPolicy"
$LW_GPO_Reg_Naam = "fMinimizeConnections"
$LW_GPO_Reg_Type = "DWORD"
$LW_GPO_Reg_Value = "0"
#endregion

#region 1.3 Qastor installatiepad
$LW_Qastor_Pad = "C:\Program Files (x86)\QPS\Qastor 2.5"
$LW_Qastor_Naam = "Qastor.exe"
$LW_Qastor = "$LW_Qastor_Pad\$LW_Qastor_Naam"
#endregion

#region 1.4 Korps huidige gebruiker
$LW_Korps = $LW_XRUpgraders | Where-Object LoodsNr -eq ([int]($env:COMPUTERNAME.Substring(4))) | Select-Object Korps
#endregion

#endregion

Start-Sleep -Seconds 10

#region 2. Uitvoering
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#region 2.1 Qastor upgrade
$LW_PopUp1 = [System.Windows.Forms.MessageBox]::Show('Dit script zal Qastor upgraden naar de nieuwste 2.5 versie. Dit neemt ongeveer 5 min. in beslag. Wil u hiermee doorgaan?', 'Upgrade Qastor', 'YesNo', 'Error')
switch ($LW_PopUp1) {
    "yes" {
        Write-Output "Qastor wordt stil gelegd en zal geupdatet worden"
        Stop-Process -Name "Qastor" -Force -ErrorAction SilentlyContinue
        Start-Process msiexec.exe -ArgumentList "/quiet /i `"$LW_Bron\$LW_Naam\Qastor\Setup Qastor 2.5.msi`"" -Wait
        Write-Output "Installatie van Qastor geslaagd"
    }
    "no" {
        Exit
    }
}
#endregion

#region 2.2 Combinatie WWAN/WLAN activeren
Write-Output "Activeren van Group Policy"
IF (!(Test-Path -Path $LW_GPO_Reg_Pad)) {
    New-Item -Path $LW_GPO_Reg_Pad
}
New-ItemProperty -Path $LW_GPO_Reg_Pad -Name $LW_GPO_Reg_Naam -PropertyType $LW_GPO_Reg_Type -Value $LW_GPO_Reg_Value -Force
#endregion

#region 2.3 Kopiëren projecten, toevoegen snelkoppeling(en)
switch ($LW_Korps.Korps) {
    "RL" { $LW_Project = @("Rivier XR2 Lite") }
    "GL" { $LW_Project = @("Kanaal XR2 Lite") }
    "SM" { $LW_Project = @("Rivier XR2 Lite") }
    "KL" { $LW_Project = @("Kust XR2 Lite") }
    "ZG" { $LW_Project = @("Rivier XR2 Lite", "Kanaal XR2 Lite") }
}
for ($i = 0; $i -lt $LW_Project.Count; $i++) {
    $LW_Bestandsnaam = "$LW_Bron\$LW_Naam\Projects\" + $LW_Project[$i]
    Copy-Item -Path $LW_Bestandsnaam -Destination $LW_Doel -Recurse
    $LW_Qastor_Link_Naam = "$env:USERPROFILE\Desktop\" + $LW_Project[$i] + ".lnk"
    $LW_Qastor_Link_Argument = "'" + "`"" + "$LW_Doel\" + $LW_Project[$i] + "\" + $LW_Project[$i] + ".qas" + "`"" + "'"
    $shell = New-Object -ComObject WScript.Shell
    $snelkoppeling = $shell.CreateShortcut("$LW_Qastor_Link_Naam")
    $snelkoppeling.Arguments = $LW_Qastor_Link_Argument
    $snelkoppeling.WorkingDirectory = $LW_Qastor_Pad
    $snelkoppeling.TargetPath = $LW_Qastor
    $snelkoppeling.save()
}
#endregion

#region 2.4 Toevoegen Wi-Fi's XR
$LW_XR_Wifi = Get-ChildItem -Path "$LW_Bron\$LW_Naam\WiFi\*.xml"
foreach ($item in $LW_XR_Wifi) {
    $LW_XR_WiFiNaam = $item.FullName
    netsh wlan add profile filename="$LW_XR_WiFiNaam"
}
#endregion

#region 2.5 Symantec Endpoint Protection verwijderen
Install-PackageProvider -Name nuget -Force
Install-Module -Name nuget -Force
Uninstall-Package -Name $LW_AntiVirus -ErrorAction Ignore
#endregion

#region 2.6 Wegschrijven checkbestand
New-Item -Path $LW_Check -Name "$LW_Naam.txt"
#endregion

#endregion

# Aanpassen softlock
# start-process -filepath "C:\Program Files\Common Files\QPS\License-Manager\license-manager.exe" -ArgumentList "-silent -activate JIVX..."