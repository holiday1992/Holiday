1.	Introduction
PST Collector is a tool which can be used to search computers in your organization to find Outlook PST files. It helps you get an inventory of PST files that are scattered throughout your organization. After you find PST files, you can use the tool to copy them in a central location which allows you to import them to Exchange Online mailboxes. After you’re confident that the PST files that you collected have been successfully imported to Office 365, you can use the tool to delete them from their original location on your network.

2.	Design Overview
•	We write this tool with Powershell script as it’s target to IT Admin or operation, they should be more comfortable with script language instead of compiled language as script language is more transparent. 
•	We should use libraries with relatively lower version to make it compatible as much as possible.
•	All actual operations should happen on client directly to reduce unexpected risks. Master machine/script will only be used for pushing task and collecting result.
•	We should cover general user cases as many as possible.

3.	Feature Overview
•	Find PST files – This feature will find all PST files on the target machines and only upload the results to a collection location. It’s designed to help admin get an inventory of the PST files.
•	Copy PST files – This feature will find all PST files on the target machines and upload both the results and PST files to a collection location. 
•	Import PST files to Office 365 – After ‘Copy PST files’, admin can use this feature to import the PST files from the collection location to Office 365.
•	Delete PST files – Once Admins are confident that the PST files have been successfully imported to Office 365, they can use this feature to delete the PST files from the original location. 

4.	Support Scenario 
This tool is designed to support general user cases as many as possible, for the special user scenario, we could also provide support based on the requirement.
These are the scenarios we support:
•	The organization has Active Directory, all machines are managed in Domain, the operator has domain admin privilege and has remote access to all the machines. In this case, the tool will work for all the machines.
•	The operator only wants to do the operation on a group of machines, if the operator has admin and remote access to the group of machines, the tool will work.
•	The operator only wants to do the operation on local machine, if the operator is the admin on the machine, the toll will work.
