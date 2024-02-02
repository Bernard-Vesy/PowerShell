
#Install-Module AzureAD
$CSVPath = "E:\tobedel\Sites - Owners.xlsx"
$line = 0

Connect-SPOService -Url https://lemo-admin.sharepoint.com
Connect-AzureAD

$excel = New-Object -ComObject Excel.Application

# Ajouter un nouveau Workbook
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

$sitelist = Get-SPOSite -Limit ALL
 
$siteownerlist = foreach($site in $sitelist){
    $ownerlist = if($site.Template -like '*group*'){

        try {
            Get-AzureADGroupOwner -ObjectId $site.GroupId | Select-Object -ExpandProperty UserPrincipalName    
        }
        catch {
            <#Do this if a terminating exception happens#>
            Write-Host "ERROR - " $site.Url  "  -  " $site.Title
        }
        
    }
    else{
        $line = $line + 1
        $worksheet.Cells.Item($line, 1) = $site.Url 
        $worksheet.Cells.Item($line, 2) = $site.Title
        $worksheet.Cells.Item($line, 3) = $site.Owner
        
    }

    [PSCustomObject]@{
        'Site Title' = $site.Title
        URL          = $site.Url
        'Owner(s)'   = $ownerlist -join '; '
        
    }
    #write-host $site.Url "|" $site.Title "|" $ownerlist 
    

    # Initialize the loop
    #'System.String'
    if ($ownerlist -is [System.String]) {

          # Output the current item
          $line = $line + 1
          $worksheet.Cells.Item($line, 1) = $site.Url 
          $worksheet.Cells.Item($line, 2) = $site.Title
          $worksheet.Cells.Item($line, 3) = $ownerlist
    }
    else {
        for ($i = 0; $i -lt $ownerlist.Length; $i++)
        {
            # Output the current item
            $line = $line + 1
                $worksheet.Cells.Item($line, 1) = $site.Url 
                $worksheet.Cells.Item($line, 2) = $site.Title
                $worksheet.Cells.Item($line, 3) = $ownerlist[$i]
            
        }
    }

    
}

  

$workbook.SaveAs($CSVPath)
$excel.Quit()
#$siteownerlist | Export-Csv -path $CSVPath -NoTypeInformation`