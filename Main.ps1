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
$form.Size = New-Object System.Drawing.Size(400,300)
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

# Create Label and TextBox for Admin Email (for adding)
$labelAdminEmail = New-Object System.Windows.Forms.Label
$labelAdminEmail.Location = New-Object System.Drawing.Point(10,100)
$labelAdminEmail.Size = New-Object System.Drawing.Size(100,20)
$labelAdminEmail.Text = "Admin Email:"

$textBoxAdminEmail = New-Object System.Windows.Forms.TextBox
$textBoxAdminEmail.Location = New-Object System.Drawing.Point(120,100)
$textBoxAdminEmail.Size = New-Object System.Drawing.Size(250,20)

# Create Label and TextBox for Admin Name (for adding)
$labelAdminName = New-Object System.Windows.Forms.Label
$labelAdminName.Location = New-Object System.Drawing.Point(10,140)
$labelAdminName.Size = New-Object System.Drawing.Size(100,20)
$labelAdminName.Text = "Admin Name:"

$textBoxAdminName = New-Object System.Windows.Forms.TextBox
$textBoxAdminName.Location = New-Object System.Drawing.Point(120,140)
$textBoxAdminName.Size = New-Object System.Drawing.Size(250,20)

# Create ComboBox for selecting Admin to Remove
$labelAdminRemove = New-Object System.Windows.Forms.Label
$labelAdminRemove.Location = New-Object System.Drawing.Point(10,180)
$labelAdminRemove.Size = New-Object System.Drawing.Size(100,20)
$labelAdminRemove.Text = "Remove Admin:"

$comboBoxAdminRemove = New-Object System.Windows.Forms.ComboBox
$comboBoxAdminRemove.Location = New-Object System.Drawing.Point(120,180)
$comboBoxAdminRemove.Size = New-Object System.Drawing.Size(250,20)

# Create Button for executing the operation
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Location = New-Object System.Drawing.Point(120,220)
$buttonExecute.Size = New-Object System.Drawing.Size(250,30)
$buttonExecute.Text = "Execute"

# Add controls to the form
$form.Controls.Add($labelApiKey)
$form.Controls.Add($textBoxApiKey)
$form.Controls.Add($labelOperation)
$form.Controls.Add($comboBoxOperation)
$form.Controls.Add($labelAdminEmail)
$form.Controls.Add($textBoxAdminEmail)
$form.Controls.Add($labelAdminName)
$form.Controls.Add($textBoxAdminName)
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
        Log-Message "Admin $email added to organization $orgName ($orgId)."
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
            Log-Message "Unauthorized access when adding admin to organization $orgName ($orgId). Skipping this organization." "ERROR"
            return
        } else {
            Log-Message "Error adding admin to organization $orgName ($orgId): $($_.Exception.Message)" "ERROR"
        }
    }
}

# Function to remove an admin
function Remove-MerakiAdmin {
    param (
        [string]$apiKey,
        [string]$orgId,
        [string]$orgName,
        [string]$adminId
    )
    
    $url = "$baseUrl/organizations/$orgId/admins/$adminId"
    
    $headers = @{
        "X-Cisco-Meraki-API-Key" = $apiKey
        "Content-Type" = "application/json"
    }
    
    try {
        $response = Invoke-RestMethod -Method Delete -Uri $url -Headers $headers
        Log-Message "Admin $adminId removed from organization $orgName ($orgId)."
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
            Log-Message "Unauthorized access when removing admin from organization $orgName ($orgId). Skipping this organization." "ERROR"
            return
        } else {
            Log-Message "Error removing admin from organization $orgName ($orgId): $($_.Exception.Message)" "ERROR"
        }
    }
}

# Load admins to remove when selecting 'Remove Admin'
$comboBoxOperation.add_SelectedIndexChanged({
    if ($comboBoxOperation.SelectedItem -eq "Remove Admin") {
        $apiKey = $textBoxApiKey.Text
        $orgs = Get-Organizations -apiKey $apiKey
        $comboBoxAdminRemove.Items.Clear()

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
                    $comboBoxAdminRemove.Items.Add("$($admin.name) ($($admin.id))")
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
    }
})

# Execute the operation when the button is clicked
$buttonExecute.Add_Click({
    $apiKey = $textBoxApiKey.Text
    $operation = $comboBoxOperation.SelectedItem
    $adminEmail = $textBoxAdminEmail.Text
    $adminName = $textBoxAdminName.Text
    $orgAccess = "full" # or customize based on your need
    $tagsAccess = @() # or customize based on your need
    
    $orgs = Get-Organizations -apiKey $apiKey
    
    foreach ($org in $orgs) {
        if ($operation -eq "Add Admin") {
            Add-MerakiAdmin -apiKey $apiKey -orgId $org.id -orgName $org.name -email $adminEmail -name $adminName -orgAccess $orgAccess -tagsAccess $tagsAccess
            Log-Message "Attempted to add admin $adminEmail to organization $($org.name)."
        }
        elseif ($operation -eq "Remove Admin") {
            $selectedAdmin = $comboBoxAdminRemove.SelectedItem
            $adminId = $selectedAdmin -replace '^.*\((.*?)\)$','$1'
            Remove-MerakiAdmin -apiKey $apiKey -orgId $org.id -orgName $org.name -adminId $adminId
            Log-Message "Attempted to remove admin $adminId from organization $($org.name)."
        }
    }
})

# Show the Form
$form.ShowDialog()
