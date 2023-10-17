# - Import le set de donnée dans une variable $DATA
#$data = Import-Excel -Path "C:\Users\bve.LEMO\LEMO SA\IS - Information Services - Documents\Digital\Project\WorkPlace\Bundles\Environnement\Environnement-POM.xlsx" 
#$data = Install-Module -Name Import-Excel -Path "E:\PLM\Ref-09-27-2023-10-28-ss.xlsx" 
#Install-Module -Name Import-Excel -Scope CurrentUser
$data = Import-Excel -Path "e:\PLM3\Ref-10-12-2023-09-12-ss.xlsx" 

#Migration des fichiers d'un serveur de fichiers vers un site SharePoint
#Connect to PNP
#Disconnect-PnPOnline
$weburl = "https://lemo.sharepoint.com/sites/LEGR-T-E"
#Connect-PnPOnline -Url $weburl -UseWebLogin
Connect-PnPOnline -Url $weburl -Interactive
$DocLib = "ArchivePLM"


foreach($line in $data)
{

    $Reference = "" 
    $Folder = $line.'CompleteFileName'

    $line.'Start Date'  = [datetime]::parseexact($line.'Start Date', 'dd.MM.yyyy HH:mm:ss', $null).ToString('MM-dd-yyyy HH:mm:ss')
     
    if (![string]::IsNullOrEmpty($line.'Folder Name'))
    {
        $Editor = $line.'AuthorEmail'
        $Author = $line.'AuthorEmail'

		$NewFileName = $files[0].PSChildName

        $DocLibPath = "/"+ $DocLib + "/" + $line.'Folder Name'
        
        $DocLibPath = "/sites/LEGR-T-E/"+ $DocLib + "/" + $line.'Folder Name'

        $Folder = Get-PnPFolder -Url $DocLibPath -Includes ListItemAllFields

        #write-host $Folder.ListItemAllFields.Id

        Set-PnPListItem -List $DocLib -Identity $Folder.ListItemAllFields.Id  -Values @{Author=$Author; Editor=$Editor; Modified=$line.'Start Date'; Created=$line.'Start Date'; Project=$line.'Etude'} 
        }
    }