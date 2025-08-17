# WindowsOsBuild

# Inventaire Windows AD -> Excel + Discord Webhook

Script PowerShell permettant de :
- Interroger Active Directory sur plusieurs OU
- Collecter la version de Windows, build, UBR, DisplayVersion
- Générer un rapport **Excel (.xlsx)** formaté
- Envoyer le rapport directement sur **Discord** via Webhook

## Prérequis
- Windows avec **PowerShell 5.1+**
- Module `ActiveDirectory`
- Module [`ImportExcel`](https://github.com/dfinke/ImportExcel)
- Accès réseau aux machines (WMI/DCOM)
- Webhook Discord valide

## Utilisation

```powershell
  -Install-Module ImportExcel -Scope CurrentUser
.\version-windows.ps1 `
  -SearchBases @("OU=Domain Controllers,DC=TEST,DC=local","OU=SRV-TEST,OU=TEST,DC=TEST,DC=local") `
  -OutFolder "C:\Temp" `
  -WebhookUrl "https://discord.com/api/webhooks/xxxxx"

