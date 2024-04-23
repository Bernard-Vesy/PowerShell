# - Import le set de donnée dans une variable $DATA
#$data = Import-Excel -Path "C:\Users\bve.LEMO\LEMO SA\IS - Information Services - Documents\Digital\Project\WorkPlace\Bundles\Environnement\Environnement-POM.xlsx" 
#$data = Install-Module -Name Import-Excel -Path "E:\PLM\Ref-09-27-2023-10-28-ss.xlsx" 
#Install-Module -Name Import-Excel -Scope CurrentUser
$data = Import-Excel -Path "C:\Users\bve\LEMO SA\IS - Information Services - Digital\SR-ER\SR_240412_005\REB_Achats.xlsx" 

#Migration des fichiers d'un serveur de fichiers vers un site SharePoint
#Connect to PNP
#Disconnect-PnPOnline
$weburl = "https://lemo.sharepoint.com/sites/LEGR-REB-TIMEPIECES"
#Connect-PnPOnline -Url $weburl -UseWebLogin
Connect-PnPOnline -Url $weburl -Interactive
$DocLib = "REB_Achats"


# - Delete the library content ----------------------------------
Get-PnPList -Identity $DocLib | Get-PnPListItem -PageSize 100 -ScriptBlock {
    Param($items) Invoke-PnPQuery } | ForEach-Object { $_.Recycle() | Out-Null
}

foreach($line in $data)
{
#    $line.'File Name '
#    $line.'File Size'
#$line.'reference'

$Reference = "" 
    #search file name in the sub directory
    #$Folder = "\\ntlemo-webfs-1-p\Portail\" + $line.'reference'
    #$Folder = $line.'CompleteFileName'
    $Folder = $line.'FullPath'

    $NumberOfFiles = 0
    if (Test-Path -Path $Folder) {
        # Path exists!
        $files=Get-ChildItem $Folder
        $NumberOfFiles = $files.Count
        if($NumberOfFiles -eq 1)
        {
            $Reference = $files[0].FullName
            $line.'Created'     = $files[0].CreationTimeUtc
            $line.LastModified =  $files[0].LastAccessTimeUtc
        }
        else {
              write-host " more than 1 file : Count = "  $NumberOfFiles " : " $Folder
        }
    } 
    else {
        write-host " folder not exist " $Folder
    }
    
     
    if (![string]::IsNullOrEmpty($Reference))
    {

        # Archived or not -----------------------------------
        #[bool]$Archived = $false

        #switch ($line.'archive')
        #    {
        #        "TRUE"  { $Archived=$true }
        #        "FALSE" { $Archived=$false }
         #   }
		
		#write-host $files[0].PSChildName
		
		
		#if($line.'referenceClean' -ne ""){
		#	$NewFileName = $line.'referenceClean'
		#}
        
        #"ArchivePLM/E4378 - REDEL 2P"
        #$DocLibPath = $DocLib + "/" + $line.'F5'
        $DocLibPath = $DocLib 


        # Replace "<BR>" by Carrege Return Line Feed
        #$desc1 = $line.'description' -replace("<BR>","`r`n")

        $Editor = 'svc_transfert_spo@lemo.com'
        $Author = 'svc_transfert_spo@lemo.com'

      
        #Add-PnPFile  -Path $Reference -Folder $DocLibPath  -NewFileName $line.'Fichier Original' -Values @{Title=$line.DescShort;_ExtendedDescription=$line.Desc;Author=$Author; Editor=$Editor; Modified=$line.'Start Date'; Created=$line.'Start Date'; Project=$line.'Etude'}
        Add-PnPFile  -Path $Reference -Folder $DocLibPath -NewFileName $files[0].Name -Values @{Author=$Author; Editor=$Editor; Modified=$line.LastModified ; Created=$line.Created; Departement=$line.F6;Year=$line.F7;Supplier=$line.F8}

        }
    }