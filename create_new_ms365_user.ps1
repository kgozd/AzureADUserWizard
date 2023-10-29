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

    }catch{
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

function retrieve_OldUserDataForNewUser {
    param($OldUserEmail)
    
    $oldUserGroupMembership = Get-AzureADUser -ObjectId $OldUserEmail | Get-AzureADUserMembership
    $oldUserManager = Get-AzureADUserManager -ObjectId $OldUserEmail
    $oldUserData = Get-AzureADUser -ObjectId $OldUserEmail
    $oldUserLicenseType = Get-AzureADUserLicenseDetail -ObjectId $OldUserEmail
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
        [int] $Length = 16  # Długość hasła (domyślnie 12)
    )

    $CharacterSet = [System.Collections.Generic.List[char]]@()
    $CharacterSet.AddRange([char[]]'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+-=[]{}|;:,.<>?')

    $SecureRandom = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
    $RandomBytes = New-Object byte[] $Length
    $SecureRandom.GetBytes($RandomBytes)

    $Password = -join ($RandomBytes | ForEach-Object { $CharacterSet[$_ % $CharacterSet.Count] })
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.ForceChangePasswordNextLogin = $false
    $PasswordProfile.EnforceChangePasswordPolicy = $false
    $PasswordProfile.Password = $Password
    return $PasswordProfile
}

function loading_animation {
    param($DurationInSeconds = 35)

    $AnimationChars = @('\', '|', '/', '-')
    $StartTime = Get-Date
    Write-Host "Configuring Exchange"
    Write-Host "This may take a while..."
    while ((Get-Date).AddSeconds(-$DurationInSeconds) -lt $StartTime) {
        foreach ($char in $AnimationChars) {
            Write-Host -NoNewline $char 
            Start-Sleep -Milliseconds 250
            [Console]::SetCursorPosition(([Console]::CursorLeft - 1), [Console]::CursorTop)
        }
    }
}

function create_newuser {
    $name_of_newuser, $surname_of_newuser = retrieve_newuser_name_and_surname
    $Email, $Smtp = create_email_address -FirstName $name_of_newuser -LastName $surname_of_newuser
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
    
    Clear-Host
    try {
        $userID = (Get-AzureADUser -ObjectId $($Email + "@" + $($oldUserData.OldUserDomain))).ObjectId
        $oldusermanagerID = (Get-AzureADUserManager -ObjectId $olduser_email).ObjectId
        Set-AzureADUserManager -ObjectId $userID -RefObjectId $oldusermanagerID
    }
    catch {
        Write-Host "No managers has been assigned!" -ForegroundColor Red
    }
    try {
        ForEach ($group in $oldusergroups) {
            Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $userID
        }
    }
    catch {
        Write-Host "No groups has been assigned!" -ForegroundColor Red
    }
    



    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $License.SkuId = $oldUserData.LicenseType   
    $assignlic = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $assignlic.AddLicenses = $License
    Set-AzureADUserLicense -ObjectId $($Email + "@" + $($oldUserData.OldUserDomain))  -AssignedLicenses $assignlic
    
    loading_animation
    Set-Mailbox $($Email + "@" + $($oldUserData.OldUserDomain)) -EmailAddresses "SMTP:$($Smtp + "@" + $($oldUserData.OldUserDomain))"
    Clear-Host

    Write-Host "User $name_of_newuser $surname_of_newuser has been created succesfully!!! `n"  -ForegroundColor Green
    Write-Host "Created user MS365 login: " -ForegroundColor DarkCyan
    Write-Host $($Email + "@" + $($oldUserData.OldUserDomain)) `n 
    Write-Host "Created user MS365 password: "  -ForegroundColor DarkCyan
    Write-Host  $PasswordProfile.Password `n`n 
    $elo = Read-Host "Do you want to see user config?(y/n): "
    if ($elo -eq "n") { 
        exit
    }
    elseif ($elo -eq "y") {
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

