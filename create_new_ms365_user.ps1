# Install-Module -Name AzureAD
# Install-Module -Name ExchangeOnlineManagement

Clear-Host
Import-Module  -Name AzureAD 


function connect_ms_apps {
    try {

        $credentials = Get-Credential -Message "Please enter your credentials for Microsoft 365"
        $password = $credentials.Password
        $username = $credentials.UserName
        $UserCredential = New-Object System.Management.Automation.PSCredential ($username, $password)
        Connect-AzureAD -Credential $UserCredential
        Connect-ExchangeOnline -Credential $UserCredential
        Clear-Host

    }
    catch {
        Write-Host "Not proper password or username provided!!!" -ForegroundColor Red
        Exit
    }
}


function close_ms_sessions {
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-AzureAD -Confirm:$false
}


function retrieve_newuser_name_and_surname {
    $name = Read-Host "Type name of a new user"
    $surname = Read-Host "Type surname of a new user"
    return $name, $surname
}


function retrieve_olduser_email_address_for_new_user {
    do {
        $oldUserEmail = Read-Host "Type an email of the old user"
        $user = Get-AzureADUser -ObjectId $oldUserEmail.Trim()
        if ($user) {
            return $user.Mail
        }
        else {
            Clear-Host
            Write-Host "User <$oldUserEmail> not found in your company!!!" -ForegroundColor Red
            $answer = Read-Host "Do you want to try again? (y/n) "
            Clear-Host
            if ($answer -eq "n") { 
                Write-Host "User creation has been cancelled!"  -ForegroundColor Red
                exit
            }
        }  
    } while ($true)
}


function select_license_type_for_new_user {
    $available_licenses = @()
    $licenses_across_organisation = Get-AzureADSubscribedSku 
    $license_number = 0

    ForEach ($license in $licenses_across_organisation ) {
        $enabled_licenses = $license | Select-Object @{Name = 'PrepaidUnitsEnabled'; Expression = { $_.PrepaidUnits.Enabled } }
        $available_units = [int]$enabled_licenses.PrepaidUnitsEnabled - [int]$license.ConsumedUnits
        # $available_units = [int]$license.ConsumedUnits
        if ($available_units -gt 0 -and $available_units -lt 1000 ) {
            $license_number += 1

            $license_info = New-Object PSObject -Property @{
                LicenseNumber        = $license_number
                NumberOfFreeLicenses = $available_units
                LicenseType          = $license.SkuPartNumber
                LicenseSkuId         = $license.SkuId
            }
            $available_licenses += $license_info
        }
    }
    $available_licenses_output = $available_licenses | Format-Table -Property LicenseNumber, LicenseType, NumberOfFreeLicenses -AutoSize | Out-String

    if ($available_licenses.Count -eq 0 ) {
        $answer = Read-Host "No available licenses. An Exchange mailbox will not be created.`nDo you want to proceed? (y/n)" 
        if ($answer -ne "y") {
            Write-Host "The account has not been created."
            Start-Sleep 5
            close_ms_sessions
            return $null
        }
    }
    else {
        do {
            $selected_number = Read-Host  -Prompt "$available_licenses_output`nChoose the license you want to assign to the user.`nType the chosen LicenseNumber"         
            $selected_license = $available_licenses | Where-Object { $_.LicenseNumber -eq $selected_number }

            if ($selected_license) {
                return  $selected_license.LicenseSkuId
            }
            else {
                Write-Host "Chosen license number is not available, please try again" -ForegroundColor Red
            }
        } while ($true)
    }
}



function retrieve_OldUserDataForNewUser {
    param($OldUserEmail)
    
    $oldUserGroupMembership = Get-AzureADUser -ObjectId $OldUserEmail | Get-AzureADUserMembership
    $oldUserManager = Get-AzureADUserManager -ObjectId $OldUserEmail
    $oldUserData = Get-AzureADUser -ObjectId $OldUserEmail
    $OldUserDomain = $OldUserEmail -split "@" | Select-Object -Last 1 

    $userInfo = [PSCustomObject]@{
        OldUserEmail  = $OldUserEmail
        OldUserDomain = $OldUserDomain
        Manager       = $oldUserManager.Mail
        UserJob       = $oldUserData.JobTitle
        UserCountry   = $oldUserData.Country
        UserLocation  = $oldUserData.UsageLocation
        LicenseType   = $oldUserLicenseType.SkuId
    }

    return $userInfo, $oldUserGroupMembership
}


function create_email_address {
    param (
        [string] $FirstName,
        [string] $LastName
    )

    function remove_diacritics {
        param (
            [string] $text
        )

        $normalized = $text.Normalize([System.Text.NormalizationForm]::FormD)
        $builder = New-Object Text.StringBuilder

        $normalized.ToCharArray() | ForEach-Object {
            if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$builder.Append($_)
            }
        }

        return $builder.ToString()
    }

    $FirstName = remove_diacritics -text $FirstName 
    $LastName = remove_diacritics -text $LastName 

    $Smtp = $FirstName + "." + $LastName
    $Smtp = $Smtp -replace '[^a-zA-Z0-9.]', ''
    $Smtp = $Smtp.ToLower() 

    $Email = $FirstName[0] + $LastName
    $Email = $Email -replace '[^a-zA-Z0-9]', ''
    $Email = $Email.ToLower() 

    return $Email, $Smtp
}


function generate_user_password {
    param (
        [int] $length = 16  # Długość hasła (domyślnie 12)
    )
    $CharacterSet = [System.Collections.Generic.List[char]]@()
    $CharacterSet.AddRange([char[]]'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+-=[]{}|;:,.<>?')

    $SecureRandom = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
    $RandomBytes = New-Object byte[] $length
    $SecureRandom.GetBytes($RandomBytes)

    $Password = -join ($RandomBytes | ForEach-Object { $CharacterSet[$_ % $CharacterSet.Count] })
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.ForceChangePasswordNextLogin = $false
    $PasswordProfile.EnforceChangePasswordPolicy = $false
    $PasswordProfile.Password = $Password
    return $PasswordProfile
}


function create_newuser {
    $name_of_newuser, $surname_of_newuser = retrieve_newuser_name_and_surname
    $Email, $Smtp = create_email_address -FirstName $name_of_newuser -LastName $surname_of_newuser
    $chosen_license_skuid = select_license_type_for_new_user 
    $olduser_email = retrieve_olduser_email_address_for_new_user
    $oldUserData, $oldusergroups = retrieve_OldUserDataForNewUser -OldUserEmail $olduser_email
    $PasswordProfile = generate_user_password


    New-AzureADUser `
        -PasswordProfile $PasswordProfile    `
        -DisplayName $($name_of_newuser + " " + "$surname_of_newuser" )     `
        -UserPrincipalName $($Email + "@" + $($oldUserData.OldUserDomain))    `
        -AccountEnabled $true       `
        -JobTitle $($oldUserData.UserJob) `
        -MailNickName $Email    `
        -GivenName $name_of_newuser   `
        -Surname   $surname_of_newuser `
        -UsageLocation  $($oldUserData.UserLocation)
    
    # Clear-Host
    try {
        $userID = (Get-AzureADUser -ObjectId $($Email + "@" + $($oldUserData.OldUserDomain))).ObjectId
        $oldusermanagerID = (Get-AzureADUserManager -ObjectId $olduser_email).ObjectId
        Set-AzureADUserManager -ObjectId $userID -RefObjectId $oldusermanagerID
    }
    catch {
        Write-Host "No managers has been assigned!" -ForegroundColor Red
    }

    #adding license to new user
    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $License.SkuId = $chosen_license_skuid
    $assignlic = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $assignlic.AddLicenses = $License
    Set-AzureADUserLicense -ObjectId $($Email + "@" + $($oldUserData.OldUserDomain))  -AssignedLicenses $assignlic

    try {
        ForEach ($group in $oldusergroups) {
            Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $userID
            Add-DistributionGroupMember -Identity  $group.ObjectId -Member  $userID
        }
    }
    catch {}
    $old_user_groupcount = $oldusergroups.Count
    $new_user_groupcount = (Get-AzureADUserMembership -ObjectId $userID).Count
    
    if ($new_user_groupcount -ne $old_user_groupcount) {
        Write-Host "The number of groups assigned to the new user is different from that of the old user."
    }

    Write-Host "Configuring Exchange"
    Write-Host "This may take a while..."
    try {
        Set-Mailbox $($using:Email + "@" + $($using:olduserdomain)) -EmailAddresses "SMTP:$($using:Smtp + "@" + $($using:olduserdomain))"
        # Set-Mailbox -Identity "GradyA@7qcgzb.onmicrosoft.com" -EmailAddresses "SMTP:grady.archi@7qcgzb.onmicrosoft.com"
    }
    catch {
        Write-Host "Exchange not configured! This may be caused by no license attached to the account." -ForegroundColor yellow
    }

    Write-Host "User $name_of_newuser $surname_of_newuser has been created succesfully!!! `n"  -ForegroundColor Green
    Write-Host "Created user MS365 login: " -ForegroundColor DarkCyan
    Write-Host $($Email + "@" + $($oldUserData.OldUserDomain)) `n 
    Write-Host "Created user MS365 password: "  -ForegroundColor DarkCyan
    Write-Host  $PasswordProfile.Password `n`n 
    $yes_or_no = Read-Host "Do you want to see user config?(y/n)"
    $yes_or_no = Read-Host "Do you want to see user config?(y/n)"
    if ($yes_or_no -eq "n") { 
        exit
    }
    elseif ($yes_or_no -eq "y") {
        Get-AzureADUser -ObjectId $($Email + "@" + $($oldUserData.OldUserDomain)) | Format-List
    }
    else {
        Write-Host "Invalid input. Please enter 'y' or 'n'." -ForegroundColor Red
    }

}


function main {
    connect_ms_apps 
    create_newuser
    close_ms_sessions
}

main 

