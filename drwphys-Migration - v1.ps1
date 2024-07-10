# - Import le set de donnée dans une variable $DATA
#$data = Import-Excel -Path "C:\Users\bve.LEMO\LEMO SA\IS - Information Services - Documents\Digital\Project\WorkPlace\Bundles\Environnement\Environnement-POM.xlsx" 
#$data = Install-Module -Name Import-Excel -Path "E:\PLM\Ref-09-27-2023-10-28-ss.xlsx" 
#Install-Module -Name Import-Excel -Scope CurrentUser
#$data = Import-Excel -Path "C:\Users\bve\LEMO SA\Project Management Tool - P-122502 - DAM - ONELEMO\Implementation\drwphys-Panel.xlsx" 

$data = Import-Excel -Path "C:\Users\bve\LEMO SA\Project Management Tool - P-122502 - DAM - ONELEMO\Implementation\drwphys.xlsx" 

#Migration des fichiers d'un serveur de fichiers vers un site SharePoint
#Connect to PNP
#Disconnect-PnPOnline
$weburl = "https://lemolab.sharepoint.com/sites/LEGR-IS-DAM"
#Connect-PnPOnline -Url $weburl -UseWebLogin
Connect-PnPOnline -Url $weburl -Interactive
$DocLib = "drwphys"
$drwphys_path = "\\194.148.36.40\Dessins_CAO_PDF\"

# - Delete the library content ----------------------------------
Get-PnPList -Identity $DocLib | Get-PnPListItem -PageSize 100 -ScriptBlock {
    Param($items) Invoke-PnPQuery } | ForEach-Object { $_.Recycle() | Out-Null
}

foreach ($line in $data) {
    #    $line.'File Name '
    #    $line.'File Size'
    #$line.'reference'

    $Reference = "" 
    #search file name in the sub directory
    #$Folder = "\\ntlemo-webfs-1-p\Portail\" + $line.'reference'
    $filePath = $drwphys_path + $line.'drwphys_name'

    if (Test-Path -Path $filePath) {
        # Action si le fichier existe
        
        # Récupération du Owner du fichiers ------------------------------------
        $Owner = Get-Acl($filePath) | Select-Object Owner | Out-String
        $Owner =  $Owner.Replace('Owner',"").Replace("-----","").Replace("LEMO\","").Replace("`r`n", "").Replace("BUILTIN\","").trim()
        if ($Owner -eq "Administrateurs") {
            $Owner = "avicario@lemo.com"
        }

        $files = Get-ChildItem -Path $filePath | Select-Object Name, Length, CreationTime, LastAccessTime, LastWriteTime, Owner 
        # Afficher les informations pour chaque fichier
        foreach ($file in $files) {
            $Creation= $($file.CreationTime)
            #Write-Output "Nom: $($file.Name)"
            #Write-Output "Taille: $($file.Length)"
            #Write-Output "Date de création: $($file.CreationTime)"
            #Write-Output "Dernier accès: $($file.LastAccessTime)"
            #Write-Output "Dernière modification: $($file.LastWriteTime)"
            #Write-Output "Owner: $($file.Owner)"
            
        }
        

        
        $Editor = $Owner
        $Author = $Owner
        #adm_bve@lemolab.onmicrosoft.com
        $Editor = "adm_bve@lemolab.onmicrosoft.com"
        $Author = "adm_bve@lemolab.onmicrosoft.com"

        $StrDte="01-01-2022 10:10:10"

        #Add-PnPFile  -Path $Reference -Folder $DocLibPath  -NewFileName $line.'Fichier Original' -Values @{Title=$line.DescShort;_ExtendedDescription=$line.Desc;Author=$Author; Editor=$Editor; Modified=$line.'Start Date'; Created=$line.'Start Date'; Project=$line.'Etude'}
        Add-PnPFile  -Path $filePath  -Folder $DocLib -Values @{Title=$line.drwphys_name;_ExtendedDescription=$line.drwphys_name;Modified=$StrDte;Created=$StrDte;Author=$Author; Editor=$Editor; drwphys_domain=$line.drwphys_domain;drwphys_id=$line.drwphys_id;drwphys_drwdetid=$line.drwphys_drwdetid;drwphys_filetype=$line.drwphys_filetype;drwphys_path=$drwphys_path;drwphys_url=$line.drwphys_url;drwphys_lang=$line.drwphys_lang.ToUpper()}
        
    }
    else {
        # Action si le fichier n'existe pas
        Write-Output  "Not Exist |$filePath"
        # Ajouter une autre action ici
    }
}