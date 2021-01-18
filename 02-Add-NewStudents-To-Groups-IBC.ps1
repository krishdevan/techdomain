##################################################################
## PART-1
## ## ==== THERE SHOULD BE P1-P6 IN THE OGV DIALOG ===
## THIS SCRIPT IMPORTS USERS FROM IBC EXCEL FILE:
## 1) CHECKS IF ALL USERS EXIST, 
## 2) ONLY ADDS EXISTING USERS TO CLASS GROUPS
## 3) FOR NON_EXISTING USERS, RUN 01-Add-NewStudentsFrom IBC.ps1
##     TO ADD THEM FIRST B4 RUNNING THIS SCRIPT
###################################################################

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

# Check if Excel File exists & if so, Import
$xlFile      = $FileBrowser.FileName	
$listOfNames = Import-Excel $xlFile -NoHeader -DataOnly

# Clean out blank lines if any
$listOfNames = $listOfNames | where {$_.P2 -inotlike "" } | OGV -PassThru

# Use this to convert into TitleCase / PascalCase later
$textinfo = (Get-Culture).TextInfo

# Array of users from the IBC
$usersArray = @()

$listOfNames | foreach {
	$nameArray = ($_.P3).Split(' ')
	$ln        = $nameArray[0].ToLower()
	$fn        = $nameArray[1].ToLower()
	$LastName  = $textinfo.ToTitleCase($ln)
	$FirstName = $textinfo.ToTitleCase($fn)
	$userNameArray = ($_.P6).Split('@')
	$samAccountName = $userNameArray[0].ToLower()
	$GreenRiverEmail = $_.P6
	$displayname     = "$FirstName $LastName"

	# Check if user already exists    
	Try {
		#Attempt to retrieve info on the user
		$user = Get-ADUser -Identity $samAccountName

		# If above line passes, the user exists.
		Out-Log "User $displayname [$samaccountName] Exists... skipping" -TextColor Yellow

		# Add this user to the users array
		$usersArray += $user

	} catch {
		#User does not exist, adding the current user
		Out-Log "User $displayname does NOT Exist..." -TextColor Cyan
	} # end catch
} #end foreach

###################################################
## PART-2
## ADDING EXISTING USERS TO SPECIFIC CLASS GROUPS 
###################################################
$continue = Read-Host "`nContinue adding users to Groups (y/n)?"
If ($continue -notmatch "y") {
	Out-Log "Exiting..." "Red"
	Exit
}

# Continue to add users to groups ...
$className = Read-Host "Class Name"

# Use my custom module to determine the next empty AD group
Find-EmptyGroupNumber -ClassName $classname

# based on the info above, select the starting group #
[int]$begin     = Read-Host "Start Group#"        
[int]$usersArrayIndex = 0
[int]$numberOfStudents = $usersArray.Count

 #add the start group number to the number of students
# to determine the end group number (sequentially)
[int]$end       = $begin + $numberOfStudents - 1

$begin..$end | foreach {
	if( ($className -eq "IT114") -or ($className -eq "IT160") )
	{ $x = "{0:D3}" -f $_ } # 3 digit class AD groups
	else {$x = "{0:D2}" -f $_} # 2 digit class AD Groups

	$groupName = $className + "_" + $x
	$userIdentity = ($usersArray[$usersArrayIndex]).SamAccountName
	$userFullName = $usersArray[$usersArrayIndex].GivenName + " " + $usersArray[$usersArrayIndex].Surname
	Out-Log "Adding $userFullName ($userIdentity) to $groupName..." -TextColor Cyan
	Add-ADGroupMember -Identity $groupName -Members $userIdentity
	
	# Reset the password if needed. Mostly not necessary. Comment out.
	#Reset-Password -changePwdAtLogon:$false -TechDomainUserName $username
	
	$usersArrayIndex++
} # end foreach



    


