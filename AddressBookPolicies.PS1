######################### Final Build ###########################
#Change Variables As Needed
$HoldingCompany = "<Company1>"
$PortfolioCompany = "<Company2>"

function New-ABP {
#Create Company Address Lists
'Creating Company Address Lists'
    #Distribution Lists
    $DLfilter = "((CustomAttribute2 -like '$HoldingCompany' -or CustomAttribute2 -like '$PortfolioCompany') -and ((ObjectCategory -like 'group') -or (RecipientType -eq 'DynamicDistributionGroup')))"
    New-AddressList -Name ($PortfolioCompany + " Distribution Lists") -RecipientFilter $DLfilter
    $DLs = Get-AddressList | Where {$_.Name -like "$PortfolioCompany Distribution Lists"}
    #Group Address List
    $Groupfilter = "(CustomAttribute2 -like '$HoldingCompany' -or CustomAttribute2 -like '$PortfolioCompany') -and (RecipientTypeDetailsValue -eq 'GroupMailbox')"
    New-AddressList -Name ($PortfolioCompany + " Groups") -RecipientFilter $Groupfilter
    $Groups = Get-AddressList | Where {$_.Name -like "$PortfolioCompany Groups"}

#Creates the Recipient Filter for Global Address List
$GALfilter = "(CustomAttribute2 -like '$HoldingCompany' -or CustomAttribute2 -like '$PortfolioCompany') -or (RecipientTypeDetails -eq 'RoomMailbox')"

#Confirm Recipient Filter Looks Correct
Get-Recipient -RecipientPreviewFilter $GALfilter | Select DisplayName, RecipientTypeDetails, CustomAttribute2

#Create the New GAL
'Creating New Global Address List'    
    New-GlobalAddressList -Name ($PortfolioCompany + " Global Address List") -RecipientFilter $GALfilter
    $GAL = Get-GlobalAddressList | Where {$_.Name -like "$PortfolioCompany*"}

#Create Offline Address Book to go With the New GAL
'Creating New Offline Address Book'
    New-OfflineAddressBook -Name ($PortfolioCompany + " Offline Address Book") -AddressLists $GAL.Name
    $OAB = Get-OfflineAddressBook | Where {$_.Name -like "$PortfolioCompany*"}


#Create Address Book Policy
'Creating Address Book Policy'
    $AddressLists = Get-AddressList | Where {$_.Name -Like "*$PortfolioCompany*"} | Select -ExpandProperty Name
    $AddressLists += "All Contacts"
    New-AddressBookPolicy -Name ($PortfolioCompany + " Address Book Policy") -AddressLists $AddressLists -RoomList "\All Rooms" -OfflineAddressBook $OAB.Name -GlobalAddressList $GAL.Name
    $ABP = Get-AddressBookPolicy -Identity ($PortfolioCompany + " Address Book Policy")
}

function Set-ABP {
#Apply Address Book Policy
'Applying Address Book Policy'
    $ABP = Get-AddressBookPolicy -Identity ($PortfolioCompany + " Address Book Policy")
    $Userfilter = "RecipientTypeDetails -eq 'UserMailbox'"
    $Users = Get-Recipient -RecipientPreviewFilter $Userfilter | Where {$_.PrimarySmtpAddress -like "*$PortfolioCompany*"}
    foreach ($User in $Users)
        {
            Write-Output "Applying Address Book Policy to $User"
            Set-Mailbox -Identity $User.DisplayName -AddressBookPolicy $ABP.Name
            Get-Mailbox -Identity $User.DisplayName | Format-List DisplayName,AddressBookPolicy
        }

$ABP | FL
}

###########################################################

function Remove-ABP {
    $DLs = Get-AddressList | Where {$_.Name -like "$PortfolioCompany Distribution Lists"}
    $Groups = Get-AddressList | Where {$_.Name -like "$PortfolioCompany Groups"}
    $GAL = Get-GlobalAddressList | Where {$_.Name -like "$PortfolioCompany*"}
    $OAB = Get-OfflineAddressBook | Where {$_.Name -like "$PortfolioCompany*"}
    $ABP = Get-AddressBookPolicy -Identity ($PortfolioCompany + " Address Book Policy")
    $Userfilter = "RecipientTypeDetails -eq 'UserMailbox'"
    $Users = Get-Recipient -RecipientPreviewFilter $Userfilter | Where {$_.PrimarySmtpAddress -like "*$PortfolioCompany*"}


foreach ($User in $Users)
    {
        Set-Mailbox -Identity $User.DisplayName -AddressBookPolicy $null
        Get-Mailbox -Identity $User.DisplayName | Format-List DisplayName,AddressBookPolicy
    }

Remove-AddressBookPolicy -Identity $ABP.Name -Confirm:$false

Remove-AddressList -Identity $Dls.Name -Confirm:$false

Remove-AddressList -Identity $Groups.Name -Confirm:$false

Remove-OfflineAddressBook -Identity $OAB.Name -Confirm:$false

Remove-GlobalAddressList -Identity $GAL.Name -Confirm:$false
}

################################################################################################################

function Set-Logic
{
    # Dependent on Recipient Type : Set CustomAttribute2
    switch($Result.RecipientTypeDetails)
        {
            UserMailbox {Set-Mailbox -Identity $Result.PrimarySMTPAddress -CustomAttribute2 $Domain}
        
            RoomMailbox {Set-Mailbox -Identity $Result.PrimarySMTPAddress -CustomAttribute2 $Domain}
            
            DynamicDistributionGroup {Set-DynamicDistributionGroup -Identity $Result.PrimarySMTPAddress -CustomAttribute2 $Domain}

            SharedMailbox {Set-Mailbox -Identity $Result.PrimarySMTPAddress -CustomAttribute2 $Domain}

            GroupMailbox {Set-UnifiedGroup -Identity $Result.PrimarySMTPAddress -CustomAttribute2 $Domain}

            MailUniversalDistributionGroup {Set-DistributionGroup -Identity $Result.PrimarySMTPAddress -CustomAttribute2 $Domain}

            MailUniversalSecurityGroup {Set-DistributionGroup -Identity $Result.PrimarySMTPAddress -CustomAttribute2 $Domain}
        }
}

function Invoke-AllSMTP
{

    Param(
        [Parameter(Mandatory=$true)]
        [string] $Domain,
        [boolean] $Affect)


$RecipientTypes = 
    {
            (RecipientTypeDetails -eq 'UserMailbox') `
        -or (RecipientTypeDetails -eq 'MailContact') `
        -or (RecipientTypeDetails -eq 'RoomMailbox') `
        -or (RecipientTypeDetails -eq 'DynamicDistributionGroup') `
        -or (RecipientTypeDetails -eq 'SharedMailbox') `
        -or (RecipientTypeDetails -eq 'GroupMailbox') `
        -or (RecipientTypeDetails -eq 'MailUniversalDistributionGroup') `
        -or (RecipientTypeDetails -eq 'MailUniversalSecurityGroup') `
        -and (Alias -ne $null) 
    }
    $Results = Get-Recipient -Filter $RecipientTypes -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -like "*@$Domain*"}

    #Only Triggers if Affect is set to True
    if($Affect -eq $true)
        {
            #Loop through each recipient with SMTP Address like $Domain
            foreach($Result in $Results)
                {
                    Set-Logic
                }
            
        }
    else
        {
            Write-Host "No Changes Were Made" -ForegroundColor Yellow
            $Results | Select Name, PrimarySMTPAddress, CustomAttribute2, RecipientTypeDetails
        }
}


