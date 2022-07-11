#partially repurposed code from https://gcits.com/knowledge-base/export-list-office-365-users-licenses-customer-tenants-delegated-administration/


#Check if MSOnline is installed 
Import-Module MSOnline
$MSOnlineInstalled = Get-Module MSOnline

If ($MSOnlineInstalled) #If MSONline already installed then run code below
{


    #Import file containing product name strings and descriptions. Must exist in same folder as script
    $ProductNames = Import-CSV "ProductNames.csv" 

    #Define report path. Currently set to local folder
    $CSVpath = "UserLicenseReport.csv" 

    #Recreate file if it already exists
    if (Test-Path $CSVPath)
    {
        Remove-Item $CSVpath
    }

    #Check with user if they have already connected to MSOL
    $AzureADConnectionStatus = Read-Host "Are you connected to MSOL? (y/n)"
    while("y","n" -notcontains $AzureADConnectionStatus )
    {
        $AzureADConnectionStatus = Read-Host "Are you connected to MSOL? (y/n)"
    }
    If ($AzureADConnectionStatus -eq 'n')
    {
        Connect-MsolService #Connect to MSOL
    }


    Write-host "Discovering users, this may take a while (approximately 10 seconds per 500 users) depending on your user selection."
    $Users = Get-MsolUser -All #Collect all users info, change filters here if you want to select only a subset of users
    


    Write-Host "Discovered " $users.Count " users."
    Write-Host "Beginning user license processing..."
    Write-Host "-------------------------------------"

    Start-Sleep -Seconds 3 #sleeping for 3 seconds for usability

    $Count = 1 #Initializing counter

    #Cycle through all users
    foreach ($User in $Users) 
    {
        
        Write-Host "Processing User $Count of " $Users.Count " : " $user.UserPrincipalName

        $Licenses = $User.Licenses #assign user licenses to the $licenses array

        #Cycle through all assigned licenses
        foreach ($License in $Licenses)
        {
            #Cycle through all products and determine if licenses assigned to the user match any SKUs
            foreach ($Product in $ProductNames)
            {
                #If match is found then generate report array and export entry to CSV
                if ($license.AccountSkuID -eq $Product.ProductSku)
                {
                    $Report = [pscustomobject][ordered]@{
                        DisplayName           = $User.DisplayName
                        UserPrincipalName     = $User.UserPrincipalName
                        AccountDisabled       = $User.BlockCredential
                        License               = $License.AccountSkuId
                        Product               = $Product.Product
                        
                    }
                    #Export array to report
                    $Report | Export-CSV -Path $CSVpath -Append -NoTypeInformation 
                }
            }
    
        }

        $Count++

    }
}
else 
{
Write-Host "MSOnline not installed. Please run 'Install-Module MSOnline' in an administrative PowerShell session to install it"

}
