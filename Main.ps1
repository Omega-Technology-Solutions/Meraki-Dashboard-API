Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
    
    $response = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
    
    return $response
}

# Function to add a new admin
function Add-MerakiAdmin {
    param (
        [string]$apiKey,
        [string]$orgId,
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
    
    $response = Invoke-RestMethod -Method Post -Uri $url -Headers $headers -Body $body
    
    return $response
}

# Function to remove an admin
function Remove-MerakiAdmin {
    param (
        [string]$apiKey,
        [string]$orgId,
        [string]$adminId
    )
    
    $url = "$baseUrl/organizations/$orgId/admins/$adminId"
    
    $headers = @{
        "X-Cisco-Meraki-API-Key" = $apiKey
        "Content-Type" = "application/json"
    }
    
    $response = Invoke-RestMethod -Method Delete -Uri $url -Headers $headers
    
    return $response
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
            $admins = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
            foreach ($admin in $admins) {
                $comboBoxAdminRemove.Items.Add("$($admin.name) ($($admin.id))")
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
            Add-MerakiAdmin -apiKey $apiKey -orgId $org.id -email $adminEmail -name $adminName -orgAccess $orgAccess -tagsAccess $tagsAccess
            [System.Windows.Forms.MessageBox]::Show("Admin added to $($org.name)")
        }
        elseif ($operation -eq "Remove Admin") {
            $selectedAdmin = $comboBoxAdminRemove.SelectedItem
            $adminId = $selectedAdmin -replace '^.*\((.*?)\)$','$1'
            Remove-MerakiAdmin -apiKey $apiKey -orgId $org.id -adminId $adminId
            [System.Windows.Forms.MessageBox]::Show("Admin removed from $($org.name)")
        }
    }
})

# Show the Form
$form.ShowDialog()
