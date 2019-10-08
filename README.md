# createNewUserandEmail
PowerShell Scripts to create new users within Alterra Tenant

## Usage
Please download all the files. Some of them depend on the other.

In the `HelperFunctions.ps1`, review the initial comment block and update the parameters for your resort. These include extension attribute 2, Active Directory Company, UPN suffix, etc. It is also critical to update the parameters to use the appropriate Exchange Server and Domain Controller.

Once the parameters in `HelperFunctions.ps1` are updated, run the scripts from the **Exchange Online PowerShell Module**. Start the module and "cd", or change directory, until you're in the directory with the scripts. Once there, run the script by typing `.\CreateNewUserAndEmail.ps1`.

If you'd like to use the CSV version to bulk create users, be sure to follow the template CSV file included in this repo. It is called `NewUsers.csv`. Either update the template to include data for your new users OR update the parameters in the `CreateNewUserAndEmail-CSV.ps1` to use CSV file in a different location.
