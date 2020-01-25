# =======================================================
# NAME: Send-ToTeams.ps1
# AUTHOR: Steven BART, StevenBart.com
# DATE: 25/01/2020
#
# VERSION 3.3
# COMMENTS_FR: Ce script permet d'envoyer des messages dans Teams, par exemple à la fin d'une séquence de tâche.
# COMMENTS_EN: This script allows you to send messages in Teams, for example at the end of a task sequence.
#
# USAGE: PowerShell.exe -ExecutionPolicy Bypass -File .\Send-ToTeams.ps1 -Status (0|1)
# =======================================================


# Paramètre -Status pour le script / Parameter -Status for the script
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Status
)

# URI du Webhook Teams / URI of the Teams Webhook
$uri = 'https://outlook.office.com/webhook/...'

# Logo du message / Logo for the message - Format PNG / GIF / JPEG
$logo = 'https://....'

# Lecture de variables de la TS / Reading TS variables
# Décommenter pour utiliser, toutes les variables disponibles ici sont disponible: https://docs.microsoft.com/fr-fr/configmgr/osd/understand/task-sequence-variables
# Uncommented to use, all the variables available here are available : https://docs.microsoft.com/en-us/configmgr/osd/understand/task-sequence-variables
#
# $SCCM_ENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
# $SMSTSCurrentActionName = $SCCM_ENV.Value("_SMSTSCurrentActionName") # Récupération de l'action problèmatique
# $SMSTSLastActionRetCode = $SCCM_ENV.Value("_SMSTSLastActionRetCode") # Récupération du code d'erreur

# Récupération via WMI, Registre ou Powershell
$DateTime = Get-Date -Format g # Date et heure
$Manufacturer = (Get-WmiObject -Class Win32_BIOS).Manufacturer # Fabricant de l'ordinateur / Computer Manufacturer
$Model = (Get-WmiObject -Class Win32_ComputerSystem).Model # Modèle de l'ordinateur/ Computer Model
[string]$SerialNumber = (Get-WmiObject win32_bios).SerialNumber # N° de série de l'ordinateur / Serial Number
$ComputerName = (Get-WmiObject -Class Win32_ComputerSystem).Name # Nom du PC de l'ordinateur / Computer Name
$IP= (Get-WmiObject win32_Networkadapterconfiguration | Where-Object{ $_.ipaddress -notlike $null }).IPaddress | Select-Object -First 1 # Adresse IP
$WinBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId # Windows Build

# Envois du message sur Teams / Send the message on Teams
function Send-ToTeams {
  Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType "application/json; charset=utf-8"
}

# $Status = 1 -> Installation réussie / Successful installation
# Création du JSON / JSON Creation
if ($Status -eq "1") {
  $body = ConvertTo-Json -Depth 4 @{
    title    = "Installation r&eacute;ussie / Successful installation"
    themeColor = "32a850"
    text   = " "
    sections = @(
      @{
        activityTitle    = 'Installation'
        activitySubtitle = "Windows 10 1909"
        activityImage    = $logo 
      },
      @{
        title = 'INFORMATION'
        facts = @(
          @{
            name  = 'Termin&eacute; &agrave; / Finished at'
            value = $DateTime
            
          },
          @{
            name  = 'Windows Build'
            value = $WinBuild
            
          },
          @{
            name  = 'Ordinateur / Computer'
            value = "$Manufacturer $Model  ($SerialNumber)"
            
          }
        )
      }
    )
  }
  Send-ToTeams # Envois à Teams / Send to Teams
} # Fin IF / End IF
 
# $Status = 0 -> Installation échouée / Failed installation
# Création du JSON / JSON Creation
if ($Status -eq "0") {
  $body = ConvertTo-Json -Depth 4 @{
    title    = "Installation &eacute;chou&eacute;e / Failed installation"
    themeColor = "ff1717"
    text   = " "
    sections = @(
      @{
        activityTitle    = 'Installation'
        activitySubtitle = "Windows 10 1909"
        activityImage    = $logo 
      },
      @{
        title = 'INFORMATION'
        facts = @(
          @{
            name  = 'Erreur à / Failed at'
            value = $DateTime
            
          },
          @{
            name  = 'Ordinateur / Computer'
            value = "$Manufacturer $Model  ($SerialNumber)"
          
          },
            @{
            name  = 'IP'
            value = $IP
            
          }
        )
      }
    )
  }
  Send-ToTeams # Envois à Teams / Send to Teams
} # Fin IF / End IF

