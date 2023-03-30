# HelloID-Task-SA-Target-ExchangeOnline-MailContactCreate
#########################################################
# Form mapping
$formObject = @{
    Name                 = $Form.Name
    DisplayName          = $form.DisplayName
    FirstName            = $form.FirstName
    Initials             = $form.Initials
    LastName             = $form.LastName
    Alias                = $form.Alias
    ExternalEmailAddress = $form.ExternalEmailAddress
}

try {
    Write-Information "Executing ExchangeOnline action: [MailContactCreate] for: [$($formObject.DisplayName)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop
    $IsConnected = $true

    $CreatedContact = New-MailContact @formobject  -ErrorAction stop

    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = "$($CreatedContact.Guid)"
        TargetDisplayName = "$($formObject.DisplayName)"
        Message           = "ExchangeOnline action: [MailContactCreate] for: [$($formObject.DisplayName)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnline action: [MailContactCreate] for: [$($formObject.DisplayName)] executed successfully"
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = "$($formobject.Name)"
        TargetDisplayName = "$($formObject.DisplayName)"
        Message           = "Could not execute ExchangeOnline action: [MailContactCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [MailContactCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
}
finally {
    if ($IsConnected){
        $exchangeSessionEnd = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
#########################################################
