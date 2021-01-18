####################################################################
## USE THIS SCRIPT TO ADD STUDENTS FROM AN IBC IMPORT
## ==== THERE SHOULD BE P1-P6 IN THE OGV DIALOG ===
### ONCE USERS ARE SUCCESSFULLY ADDED, 
##   RUN 02-ADD-NEWSTUDENTS-TO-GROUPS-IBC.PS1, TO ADD
#    THESE STUDENTS TO CLASS GROUPS SEQUENTIALLY
####################################################################
# Add the WinForm Assembly
Add-Type -AssemblyName System.Windows.Forms

$scriptDir       = $PSScriptRoot

# Open the file dialog box to select Excel file
#$ExcelFilePath = "C:\Users\kdevan\OneDrive - Green River College\VirtualMachines"
$FileBrowser   = New-Object System.Windows.Forms.OpenFileDialog -Property @{            
		#InitialDirectory  = [Environment]::GetFolderPath('Desktop')
		InitialDirectory  = $scriptDir
		Filter            = 'SpreadSheet (*.xlsx)|*.xlsx'
}
# Show the file explorer window
$null        = $FileBrowser.ShowDialog()

# Check if Excel File exists; if so Import the file
$xlFile      = $FileBrowser.FileName	
$listOfNames = Import-Excel $xlFile -NoHeader -DataOnly

# Clean out blank lines if any
$listOfNames = $listOfNames | where {$_.P2 -inotlike "" } | OGV -PassThru	

# Use this to convert into TitleCase / PascalCase later
$textinfo = (Get-Culture).TextInfo

# =============================
#  ANY LAST MINUTE BAIL OUTS?
# =============================
$continue        = Read-Host "`nContinue (y/n)?"
If ($continue -notmatch "y") {
	Out-Log "Exiting..." "Red"
	Exit
}
##################################
# Common properties of users:
$password_ss    = ConvertTo-SecureString -String 'Password01' -AsPlainText -Force
$template_obj   = Get-ADUser -Identity template
$ouPath         = 'OU=All Students,OU=Students,DC=tech,DC=div'
$exit           = 0
$skippedCount   = 0 # track existing users
$addedCount     = 0 # track new users
$newUsersAdded  = @()
$existingUsers  = @()

# When imported, the table has 6 columns - P1 to P6
$listOfNames | foreach {
	$nameArray = ($_.P3).Split(' ')
	$ln        = $nameArray[0].ToLower()
	$fn        = $nameArray[1].ToLower()
	$LastName  = $textinfo.ToTitleCase($ln)
	$FirstName = $textinfo.ToTitleCase($fn)
	$userNameArray = ($_.P6).Split('@') #Note Grade Column is empty
	$samAccountName = $userNameArray[0].ToLower()
	$GreenRiverEmail = $_.P6
	$displayname     = "$FirstName $LastName"

	# Check if user already exists    
	Try {
		#Attempt to retrieve info on the user
		$user = Get-ADUser -Identity $samAccountName

		# If above line passes, the user exists.
		Out-Log "User $samaccountName Exists... skipping" -TextColor Yellow
		$skippedCount++
		$existingUsers += $samAccountName
	} catch {
		#User does not exist, adding the current user
		Out-Log "Adding New user: $FirstName $LastName ($samAccountName)" -TextColor Cyan
		
		#Create a hash-table to 'splat' the parameters
		  $parameters                = @{
			"SamAccountName"        = $samaccountname
			"UserPrincipalName"     = $samaccountname       
			"Instance"              = $template_obj        
			"DisplayName"           = $displayname
			"GivenName"             = $FirstName
			"Name"                  = "$LastName, $FirstName"       
			"SurName"               = $LastName
			"AccountPassword"       = $password_ss
			"EmailAddress"          = $GreenRiverEmail
			"Enabled"               = $true        
			"ErrorAction"           = 'Stop'
			"Path"                  = $oupath
			"ChangePasswordAtLogon" = $false
		} # end Hashtable
		
		$ErrorActionPreference      = 'Stop'
		Try {
			New-ADUser @parameters
		} catch {
			"`nDang, Something went wrong... Exiting`n"
			Exit
		} #New-ADUser Try/catch

		#Add user to basic groups from Student Template user
		Get-ADUser -Identity template -Properties memberof | 
			select -ExpandProperty memberof | foreach {					
				Add-ADGroupMember -Identity $_ -Members $SamAccountName
			}
		$addedCount++
		$newUsersAdded += $samAccountName
		$ErrorActionPreference      = 'Continue'
		write-host "=====" -ForegroundColor green
	} # end Catch
	
} # end foreach
#
# Display useful information after adding users/groups
#
write-host "`n=====" -ForegroundColor green
Out-Log "Number of existing users $skippedCount" -TextColor Yellow
Out-Log "Existing Users List" -TextColor Yellow
Write-Output $existingUsers
Out-Log "==================================" -TextColor Cyan
Out-Log "`nNumber of users added:  $addedCount" -TextColor Green    
Out-Log "New Users List" -TextColor Green
Write-Output $newUsersAdded
   
    
