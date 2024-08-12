How to Use the Script:

- API Key Entry: Enter your Meraki Dashboard API key in the text box provided.
- Select Operation: Choose whether to "Add Admin" or "Remove Admin" using the drop-down menu.

Adding an Admin:
  - Provide the email and name of the admin you want to add.

Removing an Admin:
- Once you select "Remove Admin", the script will fetch the list of all admins across all organizations and populate the drop-down menu.
- Select the admin you wish to remove from the drop-down list.
Execute: Click the "Execute" button to perform the operation.


Additional Notes:
The script automatically loops through all organizations associated with your API key.
Ensure you have the necessary permissions to manage admins in the organizations.
The operation is executed for all organizations linked to your API key.
This script offers a GUI for better usability and allows managing Meraki Dashboard admins across multiple organizations with a single API key.
