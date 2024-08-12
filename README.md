How to Use the Script:

- API Key Entry: Enter your Meraki Dashboard API key in the text box provided.
- Select Operation: Choose whether to "Add Admin" or "Remove Admin" using the drop-down menu.

Adding an Admin:
  - Provide the email and name of the admin you want to add.

Removing an Admin:
- Once you select "Remove Admin", the script will fetch the list of all admins across all organizations and populate the drop-down menu.
- Select the admin you wish to remove from the drop-down list.
Execute: Click the "Execute" button to perform the operation.


Logging Function (Log-Message): The Log-Message function uses Write-Host to output log messages to the console. The log messages include timestamps and are categorized as "INFO" or "ERROR."

Logging Actions: The script logs the following actions:

Successful fetching of organizations and admins.
Attempt to add or remove an admin.
Logging on 401 Unauthorized, When a 401 Unauthorized error occurs, the script logs the error along with the organization name and skips further processing for that organization.
Continues Loop:
- The script uses the continue statement to skip the current iteration when a 401 Unauthorized error is encountered during the admin listing or action.
- This ensures that the script continues to process other organizations even if one or more organizations result in a 401 Unauthorized error.

Running the Script: When you run the script, you will see the log messages directly in the PowerShell console, making it easy to monitor what the script is doing and to identify any issues.

Additional Notes:
The script automatically loops through all organizations associated with your API key.
Ensure you have the necessary permissions to manage admins in the organizations.
The operation is executed for all organizations linked to your API key.
This script offers a GUI for better usability and allows managing Meraki Dashboard admins across multiple organizations with a single API key.
