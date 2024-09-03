Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to log custom messages to the console
function Log-Message {
    param (
        [string]$message,
        [string]$type = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$type] $message"
    
    Write-Host $logMessage
}

# Create the Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Meraki Dashboard Admin Manager"
$form.Size = New-Object System.Drawing.Size(400,400)
$form.StartPosition = "CenterScreen"

# Create Label and TextBox for API Key
$labelApiKey = New-Object System.Windows.Forms.Label
$labelApiKey.Location = New-Object System.Drawing.Point(10,20)
$labelApiKey.Size = New-Object System.Drawing.Size(100,20)
$labelApiKey.Text = "API Key:"

$textBoxApiKey = New-Object System.Windows.Forms.TextBox
$textBoxApiKey.Location = New-Object System.Drawing.Point(120,20)
$textBoxApiKey.Size = New-Object System.Drawing.Size(250,20)
$textBoxApiKey.UseSystemPasswordChar = $true

# Create ComboBox for Operation Selection
$labelOperation = New-Object System.Windows.Forms.Label
$labelOperation.Location = New-Object System.Drawing.Point(10,60)
$labelOperation.Size = New-Object System.Drawing.Size(100,20)
$labelOperation.Text = "Operation:"

$comboBoxOperation = New-Object System.Windows.Forms.ComboBox
$comboBoxOperation.Location = New-Object System.Drawing.Point(120,60)
$comboBoxOperation.Size = New-Object System.Drawing.Size(250,20)
$comboBoxOperation.Items.AddRange(@("Add Admin","Remove Admin"))
$comboBoxOperation.Enabled = $false

# Create Label and TextBox for Admin Email (for adding)
$labelAdminEmail = New-Object System.Windows.Forms.Label
$labelAdminEmail.Location = New-Object System.Drawing.Point(10,100)
$labelAdminEmail.Size = New-Object System.Drawing.Size(100,20)
$labelAdminEmail.Text = "Admin Email:"

$textBoxAdminEmail = New-Object System.Windows.Forms.TextBox
$textBoxAdminEmail.Location = New-Object System.Drawing.Point(120,100)
$textBoxAdminEmail.Size = New-Object System.Drawing.Size(250,20)
$textBoxAdminEmail.Enabled = $false

# Create Label and TextBox for Admin Name (for adding - optional for display purposes)
$labelAdminName = New-Object System.Windows.Forms.Label
$labelAdminName.Location = New-Object System.Drawing.Point(10,140)
$labelAdminName.Size = New-Object System.Drawing.Size(100,20)
$labelAdminName.Text = "Admin Name:"

$textBoxAdminName = New-Object System.Windows.Forms.TextBox
$textBoxAdminName.Location = New-Object System.Drawing.Point(120,140)
$textBoxAdminName.Size = New-Object System.Drawing.Size(250,20)
$textBoxAdminName.Enabled = $false

# Create ComboBox for selecting Admin Access Level (Full Access or Read-Only)
$labelAdminAccess = New-Object System.Windows.Forms.Label
$labelAdminAccess.Location = New-Object System.Drawing.Point(10,180)
$labelAdminAccess.Size = New-Object System.Drawing.Size(100,20)
$labelAdminAccess.Text = "Admin Access:"

$comboBoxAdminAccess = New-Object System.Windows.Forms.ComboBox
$comboBoxAdminAccess.Location = New-Object System.Drawing.Point(120,180)
$comboBoxAdminAccess.Size = New-Object System.Drawing.Size(250,20)
$comboBoxAdminAccess.Items.AddRange(@("Full Access","Read-Only"))
$comboBoxAdminAccess.Enabled = $false

# Create ComboBox for selecting Organization to Add Admin
$labelOrganization = New-Object System.Windows.Forms.Label
$labelOrganization.Location = New-Object System.Drawing.Point(10,220)
$labelOrganization.Size = New-Object System.Drawing.Size(100,20)
$labelOrganization.Text = "Organization:"

$comboBoxOrganization = New-Object System.Windows.Forms.ComboBox
$comboBoxOrganization.Location = New-Object System.Drawing.Point(120,220)
$comboBoxOrganization.Size = New-Object System.Drawing.Size(250,20)
$comboBoxOrganization.Items.Add("All Organizations")
$comboBoxOrganization.Enabled = $false

# Create ComboBox for selecting Admin to Remove
$labelAdminRemove = New-Object System.Windows.Forms.Label
$labelAdminRemove.Location = New-Object System.Drawing.Point(10,260)
$labelAdminRemove.Size = New-Object System.Drawing.Size(100,20)
$labelAdminRemove.Text = "Remove Admin:"

$comboBoxAdminRemove = New-Object System.Windows.Forms.ComboBox
$comboBoxAdminRemove.Location = New-Object System.Drawing.Point(120,260)
$comboBoxAdminRemove.Size = New-Object System.Drawing.Size(250,20)
$comboBoxAdminRemove.Enabled = $false

# Create Button for executing the operation
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Location = New-Object System.Drawing.Point(120,300)
$buttonExecute.Size = New-Object System.Drawing.Size(250,30)
$buttonExecute.Text = "Execute"
$buttonExecute.Enabled = $false

# Add controls to the form
$form.Controls.Add($labelApiKey)
$form.Controls.Add($textBoxApiKey)
$form.Controls.Add($labelOperation)
$form.Controls.Add($comboBoxOperation)
$form.Controls.Add($labelAdminEmail)
$form.Controls.Add($textBoxAdminEmail)
$form.Controls.Add($labelAdminName)
$form.Controls.Add($textBoxAdminName)
$form.Controls.Add($labelAdminAccess)
$form.Controls.Add($comboBoxAdminAccess)
$form.Controls.Add($labelOrganization)
$form.Controls.Add($comboBoxOrganization)
$form.Controls.Add($labelAdminRemove)
$form.Controls.Add($comboBoxAdminRemove)
$form.Controls.Add($buttonExecute)

# Define the base URL for the Meraki API
$baseUrl = "https://api.meraki.com/api/v1"

# Function to get a list of all organizations
function Get-Organizations {
    param (
        [string]$apiKey
    )
    
    $url = "$baseUrl/organizations"
    
    $headers = @{
        "X-Cisco-Meraki-API-Key" = $apiKey
        "Content-Type" = "application/json"
    }
    
    try {
        $response = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
        Log-Message "Fetched organizations successfully."
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
            Log-Message "Unauthorized access - check API key." "ERROR"
        } else {
            Log-Message "Error fetching organizations: $($_.Exception.Message)" "ERROR"
        }
        return $null
    }
    
    return $response
}

# Function to add a new admin
function Add-MerakiAdmin {
    param (
        [string]$apiKey,
        [string]$orgId,
        [string]$orgName,
        [string]$email,
        [string]$name,
        [string]$orgAccess,
        [array]$tagsAccess
    )
    
    $url = "$baseUrl/organizations/$orgId/admins"
    
    $body = @{
        "email" = $email
        "name" = $name
        "orgAccess" = $orgAccess
        "tags" = $tagsAccess
    } | ConvertTo-Json
    
    $headers = @{
        "X-Cisco-Meraki-API-Key" = $apiKey
        "Content-Type" = "application/json"
    }
    
    try {
        $response = Invoke-RestMethod -Method Post -Uri $url -Headers $headers -Body $body
        Log-Message "Admin $email added to organization $orgName ($orgId) with $orgAccess access."
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
            Log-Message "Unauthorized access when adding admin to organization $orgName ($orgId). Skipping this organization." "ERROR"
            return
        } else {
            Log-Message "Error adding admin to organization $orgName ($orgId): $($_.Exception.Message)" "ERROR"
        }
    }
}

# Function to remove an admin based on email
function Remove-MerakiAdmin {
    param (
        [string]$apiKey,
        [string]$orgId,
        [string]$orgName,
        [string]$adminEmail
    )
    
    $url = "$baseUrl/organizations/$orgId/admins"
    
    $headers = @{
        "X-Cisco-Meraki-API-Key" = $apiKey
        "Content-Type" = "application/json"
    }
    
    try {
        # Fetch the list of admins for the organization
        $admins = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
        
        # Find the admin by email
        $admin = $admins | Where-Object { $_.email -eq $adminEmail }
        
        if ($admin) {
            $adminId = $admin.id
            
            # Remove the admin using the correct admin ID
            $urlRemove = "$baseUrl/organizations/$orgId/admins/$adminId"
            Invoke-RestMethod -Method Delete -Uri $urlRemove -Headers $headers
            Log-Message "Admin with email $adminEmail removed from organization $orgName ($orgId)."
        } else {
            Log-Message "Admin with email $adminEmail not found in organization $orgName ($orgId)." "ERROR"
        }
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
            Log-Message "Unauthorized access when removing admin from organization $orgName ($orgId). Skipping this organization." "ERROR"
            return
        } else {
            Log-Message "Error removing admin from organization $orgName ($orgId): $($_.Exception.Message)" "ERROR"
        }
    }
}

# Enable the operation dropdown when an API key is entered
$textBoxApiKey.add_TextChanged({
    if ($textBoxApiKey.Text.Length -gt 0) {
        $comboBoxOperation.Enabled = $true
        $comboBoxOrganization.Enabled = $true
        $comboBoxOrganization.Items.Clear()
        $comboBoxOrganization.Items.Add("All Organizations")

        $apiKey = $textBoxApiKey.Text
        $orgs = Get-Organizations -apiKey $apiKey

        foreach ($org in $orgs) {
            $comboBoxOrganization.Items.Add($org.name)
        }
    } else {
        $comboBoxOperation.Enabled = $false
        $textBoxAdminEmail.Enabled = $false
        $textBoxAdminName.Enabled = $false
        $comboBoxAdminAccess.Enabled = $false
        $comboBoxAdminRemove.Enabled = $false
        $comboBoxOrganization.Enabled = $false
        $buttonExecute.Enabled = $false
    }
})

# Load admins to remove when selecting 'Remove Admin'
$comboBoxOperation.add_SelectedIndexChanged({
    if ($comboBoxOperation.SelectedItem -eq "Remove Admin") {
        $textBoxAdminEmail.Enabled = $false
        $textBoxAdminName.Enabled = $false
        $comboBoxAdminAccess.Enabled = $false
        $comboBoxAdminRemove.Enabled = $true
        $comboBoxOrganization.Enabled = $false
        $buttonExecute.Enabled = $true

        $apiKey = $textBoxApiKey.Text
        $orgs = Get-Organizations -apiKey $apiKey
        $comboBoxAdminRemove.Items.Clear()
        $existingAdmins = @()
        $existingEmails = @()

        foreach ($org in $orgs) {
            $url = "$baseUrl/organizations/$($org.id)/admins"
            $headers = @{
                "X-Cisco-Meraki-API-Key" = $apiKey
                "Content-Type" = "application/json"
            }
            try {
                $admins = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
                Log-Message "Fetched admins for organization $($org.name) ($($org.id))."
                foreach ($admin in $admins) {
                    $adminEntry = "$($admin.name) - $($admin.email)"
                    if ($existingEmails -notcontains $admin.email) {
                        $existingAdmins += $adminEntry
                        $existingEmails += $admin.email
                    }
                }
            } catch {
                if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
                    Log-Message "Unauthorized access when fetching admins for organization $($org.name) ($($org.id)). Skipping this organization." "ERROR"
                    continue
                } else {
                    Log-Message "Error fetching admins for organization $($org.name) ($($org.id)): $($_.Exception.Message)" "ERROR"
                }
            }
        }

        # Sort the admins alphabetically
        $existingAdmins = $existingAdmins | Sort-Object

        # Add sorted admins to the dropdown
        foreach ($adminEntry in $existingAdmins) {
            $comboBoxAdminRemove.Items.Add($adminEntry)
        }
    } elseif ($comboBoxOperation.SelectedItem -eq "Add Admin") {
        $textBoxAdminEmail.Enabled = $true
        $textBoxAdminName.Enabled = $true
        $comboBoxAdminAccess.Enabled = $true
        $comboBoxOrganization.Enabled = $true
        $comboBoxAdminRemove.Enabled = $false
        $buttonExecute.Enabled = $true
    }
})

# Execute the operation when the button is clicked
$buttonExecute.Add_Click({
    $apiKey = $textBoxApiKey.Text
    $operation = $comboBoxOperation.SelectedItem
    $adminEmail = $textBoxAdminEmail.Text
    $adminName = $textBoxAdminName.Text
    $orgAccess = if ($comboBoxAdminAccess.SelectedItem -eq "Full Access") { "full" } else { "read-only" }
    $tagsAccess = @() # or customize based on your need
    
    if ($operation -eq "Add Admin") {
        $orgSelection = $comboBoxOrganization.SelectedItem
        $orgs = Get-Organizations -apiKey $apiKey
        
        if ($orgSelection -eq "All Organizations") {
            foreach ($org in $orgs) {
                Add-MerakiAdmin -apiKey $apiKey -orgId $org.id -orgName $org.name -email $adminEmail -name $adminName -orgAccess $orgAccess -tagsAccess $tagsAccess
            }
        } else {
            $org = $orgs | Where-Object { $_.name -eq $orgSelection }
            if ($org) {
                Add-MerakiAdmin -apiKey $apiKey -orgId $org.id -orgName $org.name -email $adminEmail -name $adminName -orgAccess $orgAccess -tagsAccess $tagsAccess
            }
        }
    }
    elseif ($operation -eq "Remove Admin") {
        $orgs = Get-Organizations -apiKey $apiKey
        $selectedAdmin = $comboBoxAdminRemove.SelectedItem
        if ($selectedAdmin) {
            $adminDetails = $selectedAdmin -split ' - '
            $adminEmailToRemove = $adminDetails[1]

            foreach ($org in $orgs) {
                Remove-MerakiAdmin -apiKey $apiKey -orgId $org.id -orgName $org.name -adminEmail $adminEmailToRemove
            }
        } else {
            Log-Message "No admin selected for removal." "ERROR"
        }
    }
})

# Show the Form
$form.ShowDialog()
