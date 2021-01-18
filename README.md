# Techdomain
This repository has scripts I wrote to help manage Tech.Div domain Identity and Access Management

1. Use the IBC-Import-Excel-Template.xlsx file as a template to copy/paste student names from 
    the IBC (instructor Brief Case). This sample file has dummy values for obvious reasons.
    NOTE: Do not copy/paste the header line from the IBC, just the data. Therefore you may end up with
    a column that is blank intentionally. The script needs this blank column, so don't delete it!!
    
2. Use the 01-Add-NewStudentFromIBC.ps1 to import the xlsx file from step 1 and create the user names
   in Tech domain AD. The user names will be the same as the GRC email prefix
   The script will check if a user already has an account and will skip creating a new account if so.
   
3. Use the 02-Add-NewStudents-To-Groups-IBC.ps1 to enforce RBAC (Role Based Access Control) to assign
   students and their respective VMs to a group which has individual control on what they can do with 
   their VMs.
