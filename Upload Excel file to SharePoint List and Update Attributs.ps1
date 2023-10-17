# - Import le set de donnée dans une variable $DATA
#$data = Import-Excel -Path "C:\Users\bve.LEMO\LEMO SA\IS - Information Services - Documents\Digital\Project\WorkPlace\Bundles\Environnement\Environnement-POM.xlsx" 
#$data = Install-Module -Name Import-Excel -Path "E:\PLM\Ref-09-27-2023-10-28-ss.xlsx" 
#Install-Module -Name Import-Excel -Scope CurrentUser
$data = Import-Excel -Path "E:\PLM\Ref-10-12-2023-09-12-ss.xlsx" 

#Migration des fichiers d'un serveur de fichiers vers un site SharePoint
#Connect to PNP
#Disconnect-PnPOnline
$weburl = "https://lemo.sharepoint.com/sites/LEGR-T-E"
#Connect-PnPOnline -Url $weburl -UseWebLogin
Connect-PnPOnline -Url $weburl -Interactive
$DocLib = "ArchivePLM"

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
    $Folder = $line.'CompleteFileName'

    #$line.'Start Date' =  Get-ChildItem -Path $line.'FullPath' | select CreationTime | Out-String 
    #$line.'Start Date' = $line.'Created'.Trim().Substring(42,19)  
    #$line.'Start Date' = [datetime]::parseexact($line.'Created', 'dd.MM.yyyy HH:mm:ss', $null).ToString('MM-dd-yyyy HH:mm:ss')

    $line.'Start Date'  = [datetime]::parseexact($line.'Start Date', 'dd.MM.yyyy HH:mm:ss', $null).ToString('MM-dd-yyyy HH:mm:ss')

    #$line.'ModificationDate' =  Get-ChildItem -Path $line.'FullPath' | select CreationTime | Out-String 
    #$line.'ModificationDate' = $line.'ModificationDate'.Trim().Substring(42,19)  
    #$line.'ModificationDate' = [datetime]::parseexact($line.'ModificationDate', 'dd.MM.yyyy HH:mm:ss', $null).ToString('MM-dd-yyyy HH:mm:ss')
    
    #$Folder = "\\DCLEMO\LE_Environnement\Conformite Fournisseurs - Copie\AA-NEW -MATERIALS SDS PLASTIC\POM\" + $line.'File Name '
    $NumberOfFiles = 0
    if (Test-Path -Path $Folder) {
        # Path exists!
        $files=Get-ChildItem $Folder
        $NumberOfFiles = $files.Count
        if($NumberOfFiles -eq 1)
        {
            $Reference = $files[0].FullName
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
        $Editor = $line.'AuthorEmail'
        $Author = $line.'AuthorEmail'
        #$Editor = "avicario@lemo.com" 
        #$Author = "lburdet@lemo.com"
        

        # Archived or not -----------------------------------
        #[bool]$Archived = $false

        #switch ($line.'archive')
        #    {
        #        "TRUE"  { $Archived=$true }
        #        "FALSE" { $Archived=$false }
         #   }
		
		#write-host $files[0].PSChildName
		$NewFileName = $files[0].PSChildName
		
		#if($line.'referenceClean' -ne ""){
		#	$NewFileName = $line.'referenceClean'
		#}

        $DocLibPath = $DocLib + "/" + $line.'Folder Name'

        # Replace "<BR>" by Carrege Return Line Feed
        #$desc1 = $line.'description' -replace("<BR>","`r`n")

       #Add-PnPFile -Path $Reference -Folder $DocLib -NewFileName $line.'Fichier Original' -Values @{"Author"=$Editor.Mail; "Editor"=$Editor.Mail;"Created"="01.01.2020 13:00:00"; "Modified"="01.01.2020 13:00:00"}
        #Add-PnPFile -Path $Reference -Folder $DocLib  -NewFileName $line.'Fichier Original' -Values @{Title=$line.DescShort;_ExtendedDescription=$line.Desc;Author=$Author; Editor=$Editor; Modified=$line.'Start Date'; Created=$line.'Start Date'}
        Add-PnPFile -Path $Reference -Folder $DocLibPath  -NewFileName $line.'Fichier Original' -Values @{Title=$line.DescShort;_ExtendedDescription=$line.Desc;Author=$Author; Editor=$Editor; Modified=$line.'Start Date'; Created=$line.'Start Date'; Project=$line.'Etude'}

        }
    }