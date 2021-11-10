#partially repurposed code from https://gcits.com/knowledge-base/export-list-office-365-users-licenses-customer-tenants-delegated-administration/


$ProductNames = Import-CSV "ProductNames.csv" #Import file containing product name strings and descriptions. Must exist in same folder as script
$CSVpath = "UserLicenseReport.csv" #Define report path. Currently set to local folder

#Recreate file if it already exists
if (Test-Path $CSVPath)
{
    Remove-Item $CSVpath
}

Connect-MsolService #Connect to MSOL

$Users = Get-MsolUser -All #Collect all users info

#Cycle through all users
foreach ($User in $Users) {
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

}