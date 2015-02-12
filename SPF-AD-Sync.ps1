#*===============================================================================
# Filename : SPFoundation-AD-Sync.ps1
# Version = "1.2.0"
#*===============================================================================
# Author = "Florent CHAUVIN"
# Company: LINKBYNET
#*===============================================================================
# Created: FCH - 12 december 2014
# Modified: FCH - 11 february 2015
#*===============================================================================
# Description :
# Script to synchronize SharePoint Foundation user profile with their domain's account
# Advanced synchronization need 'Active Directory module for Windows PowerShell' feature
# available with Windows 2008 R2 or higher.
#*===============================================================================

#*===============================================================================
# Variables Configuration
#*===============================================================================

#Path for script logging
$Global:Log = ".\" + (get-date -uformat '%Y%m%d-%H%M') + "-SPFoundation_AD_Sync.log"
#Debug mode, use to understand why account don't synchronize properly.
$Global:DebugMode = $False
#Delete account with domain unreachable or not found in domain (Advanced synchronization). The deletion is performed only if the number of account to delete is less than 30% of the number of synchronized account
$Global:DeleteUSersNotFound = $True
#Enable sending EMail
$Global:SendMail = $False
#Multiple recipients must be comma separated
$Global:emailFrom = gwmi Win32_ComputerSystem| %{$_.DNSHostName + '@' + $_.Domain}
$Global:emailTo = ""
$Global:emailCC =""
$Global:emailOnErrorTO = ""
$Global:emailOnErrorCC = ""
$Global:smtpServer = ""

#*===============================================================================
# Functions
#*===============================================================================

# Region : Create Folder if doesn't exist
function Test-FilePath-Create

{
	param([String]$FullFilename)
	If ($FullFilename -ne $null)
	{
		If (($FullFilename.substring(($FullFilename.length)-1,1)) -eq "`"")
		{
			$PathFilename = ($FullFilename.substring(0,$FullFilename.LastIndexOf("\")) + "`"")
		}
		Else
		{
			$PathFilename = ($FullFilename.substring(0,$FullFilename.LastIndexOf("\")))
		}
		If (!(Test-Path -literalPath ($PathFilename)))
		{
			New-Item $PathFilename -type directory  -errorAction SilentlyContinue | out-null
			If (Test-Path -literalPath ($PathFilename)){Write-Host "|-> Creation of folder " $PathFilename -Fore Green}
			Else {Write-Host "|-> Failed to create folder " $PathFilename -Fore Red}
		}
	}
}
#EndRegion

#Region : Load the SharePoint snap-in for PowerShell
function Load-Snapin 
{	
	<#
	
	To avoid introducing memory-leaks in your PowerShell sessions that you spawn up without using the Sharepoint Management Shell, remembe to either call SharePoint.ps1 or at least set $Host.Runspace.ThreadOptions = "ReuseThread" before executing any code.
	
	http://andersrask.sharepointspace.com/Lists/Posts/Post.aspx?ID=4
	
	#>
	
	$ver = $host | select version
	if ($ver.Version.Major -gt 1)
	{
		$Host.Runspace.ThreadOptions = "ReuseThread"
	}
 
	$snapin = (Get-PSSnapin -name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) 

	if ($snapin -ne $null) {
		Write-Host "|--> SharePoint Snap-in is loaded"	-fore Green
	} 
	else 
	{
		try 
		{
			Write-host "|--> SharePoint Snap-in not found. Action: Loading SharePoint Snap-in."
			Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
		} 
		catch
		{
			$errText = $error[0].Exception.Message
			Write-Host "|--> Loading of SharePoint Snap-in failed. Reason: $errText" -fore Red
			Exit
		}	
	}	
}
#EndRegion

#LoadActiveDirectoryModule : Load Active Directory module
function LoadActiveDirectoryModule
{
	If((($([System.Environment]::OSVersion.Version).Major -eq 6) -and ($([System.Environment]::OSVersion.Version).Minor -ge 1)) -or (([System.Environment]::OSVersion.Version).Major -gt 6))
	{
		Import-Module ServerManager
		$RSATADPowershell = Get-WindowsFeature | ?{$_.name -eq "RSAT-AD-Powershell"}
		If ($RSATADPowershell -ne $null)
		{
			if($RSATADPowershell.Installed -eq $True)
			{
				try
				{
					Import-Module ActiveDirectory
					Write-Host "|--> ActiveDirectory Module has been imported." -fore Green
					$Global:ImportModuleAD = $True
				}
				catch
				{
					$errText = $error[0].Exception.Message
					Write-Host "|--> Import of Active Directory module failed. Reason: $errText" -fore Red
					$Global:ImportModuleAD = $False
				}
				
			}
			Else
			{
				Write-Host "|--> Cannot load Active directory module because 'Active Directory module for Windows PowerShell' feature is not installed. Extended attributes won't be synchronized." -fore Red
				$Global:ImportModuleAD = $False
			}
		}
		Else
		{
				Write-Host "|--> Cannot load Active directory module because 'Active Directory module for Windows PowerShell' feature is not available on this operating system. Extended attributes won't be synchronized." -fore Red
				$Global:ImportModuleAD = $False
		}
	}
	Else
	{
			Write-Host "|--> Cannot load Active directory module because 'Active Directory module for Windows PowerShell' feature is not available on this operating system. Extended attributes won't be synchronized." -fore Red
			$Global:ImportModuleAD = $False
	}	
}
#EndRegion

#Region Test if user's domain is reachable, one time by domain for all users to synchronize
Function TestDomainAvailability
{
	Param
	(
		$_DomainName
	)
	If(($DomainReachable | Where-Object {$_ -eq $_DomainName}) -ne $null)
	{
		Write-host "  |--> The domain controller has been be listed in the previous test" -fore Green
		$Global:DomainTestedWithSuccess = $True					
	}
	ElseIf(($DomainUnReachable | Where-Object {$_ -eq $_DomainName}) -ne $null)
	{
		Write-host "  |--> The domain controller hasn't been be listed in the previous test" -fore Red
		$Global:DomainTestedWithSuccess = $False
		$Global:CounterUsersDomainUnreachable++
		$Global:UsersWithDomainUnreachable += [String]$SPuser.LoginName					
	}
	Else
	{
		# Try to list domain controller of this domain with nltest.exe
		$Nltestexe = "nltest.exe"
		$NltestParam = "/dcList:" + $_DomainName
		$NlTestResult = [String](& $nltestexe $NltestParam 2>&1)
		if($NLTestResult -match ".*ERROR.*|.*UNAVAILABLE.*")
		{
			Write-Host "  |--> Cannot list domain controller for domain $_DomainName. User synchronization can not be performed. Reason: $NlTestResult" -fore Red
			$Global:CounterUsersDomainUnreachable++
			$Global:UsersWithDomainUnreachable += [String]$SPuser.LoginName
			$Global:DomainUnReachable += $_DomainName
			$Global:DomainTestedWithSuccess = $False
		}
		ElseIf([string]::IsNullOrEmpty($NLTestResult))
		{
			Write-Host "  |--> Cannot list domain controller for domain $_DomainName. User synchronization can not be performed. Reason: Command nltest.exe send empty result, relaunch the script in new Powershell session" -fore Red
			$Global:CounterUsersDomainUnreachable++
			$Global:UsersWithDomainUnreachable += [String]$SPuser.LoginName
			$Global:DomainTestedWithSuccess = $False						
		}
		Else
		{
			Write-host "  |--> The domain controller for domain '$_DomainName' are available " -fore Green
			$Global:DomainReachable += $_DomainName
			$Global:DomainTestedWithSuccess = $True						
		}					
	}
}
#EndRegion

#Region Retrieve the user and his properties (Domain Name, SAM Account Name, SID) based on the authentication type of web application 
Function Retrieve-User-And-Properties
{
	Param
        (
			$_User
		)
	Try
	{
		if ($site.WebApplication.UseClaimsAuthentication) 
		{
			$claim = New-SPClaimsPrincipal $_User.LoginName -IdentityType WindowsSamAccountName
			$Global:SPuser  = $web | Get-SPUser -Identity $claim -ErrorAction Stop
		}
		else
		{
			$Global:SPuser = $web | Get-SPUser -Identity $_User.LoginName -ErrorAction Stop
		}
		If($claim)
		{
			[String]$SPUserStr = $Claim.value
			If ($DebugMode -eq $True)
			{
				Write-Host "Claim.value: "$Claim.value
				Write-Host "SPUserStr: "$SPUserStr
			}
		}
		Else
		{
			[String]$Global:SPUserStr = $SPUser
			If ($DebugMode -eq $True)
			{
				Write-Host "SPuser: "$SPuser
				Write-Host "SPUserStr: "$SPUserStr
			}
		}
		#Parse account name to get user name and domain
		$SplitSPuser = $SPUserStr.split("\")
		$Global:SPUserSAMAccountName = $SplitSPuser[1]
		$Global:DomainName = $SplitSPuser[0]
		If ($DebugMode -eq $True)
		{
			Write-Host "SplitSPuser: "$SplitSPuser
			Write-Host "SPUserSAMAccountName: "$SPUserSAMAccountName
			Write-Host "DomainName: "$DomainName
		}					
		If($DomainName -match "\|")
		{
			$SplitDomainName = $DomainName.split("|")
			$Global:DomainName = $SplitDomainName[1]
			If ($DebugMode -eq $True)
			{
				Write-Host "SplitDomainName: "$SplitDomainName
				Write-Host "DomainName: "$DomainName
			}					
		}
		$Global:SPUserSID = $SPUser.SID
	}
	Catch
	{
		$Global:SPuser = $null
		$errText = $error[0].Exception.Message
		Write-Host "  |--> Failed to retrieve SharePoint User and his properties. Reason: $errText" -fore Red	
	}
}
#EndRegion

#Region Get AD account and launch check for modification and update
Function GetAndCheckADAccountModification
{
	Param
	(
		$_SPUserSAMAccountName,
		$_SPUserSID,
		$_SPuser,
		$_SPUserStr,
		$_DomainName
	)

	Try
	{	
		Write-Host "  |--> Get user information from domain"
		
		#Two requests, one by SID and one by SAM Account Name to verify if account have been deleted, recreated (New SID) or modified (New SAM Account Name).
		$ADUserBySAMAccountName = get-aduser -f {SAMAccountName -eq $_SPUserSAMAccountName} -server $DomainName -properties DisplayName, EmailAddress, Department, Title, SAMAccountName, OfficePhone, MobilePhone
		
		$ADUserBySID = get-aduser -f {SID -eq $_SPUserSID} -server $DomainName -properties DisplayName, EmailAddress, Department, Title, SAMAccountName, OfficePhone, MobilePhone

		If ($DebugMode -eq $True)
		{						
			Write-Host "AD User By SAMAccountName (user properties) :" ($ADUserBySAMAccountName | select *)
			Write-Host "AD User By SID (user properties) :" ($ADUserBySID | select *)								
		}
		
		If(($ADUserBySAMAccountName -eq $null) -and ($ADUserBySID -eq $null))
		{
			Write-Host "  |--> User $SPUserSAMAccountName not found in domain $DomainName" -fore Red
			$Global:CounterUsersAdvancedNotFound++
			$Global:UsersNotFound += [String]$SPuser.LoginName
		}
		Else
		{
			CheckADAccountModificationAndUpdate -_ADUserBySAMAccountName $ADUserBySAMAccountName -_ADUserBySID $ADUserBySID -_SPuser $_SPuser -_SPUserStr $_SPUserStr -_SPUserSID $_SPUserSID -_DomainName $_DomainName
		}
	}
	Catch
	{
		$errText = $error[0].Exception.Message
		Write-Host "  |--> Cannot get user information from domain. Reason: $errText " -fore Red
		$Global:CounterUsersNativeSynchronizationFailed++
		$Global:UsersWithNativeSynchonizationError += [String]$SPuser.LoginName										
		$Global:CounterUsersAdvancedSynchronizationFailed++
		$Global:UsersWithAdvancedSynchonizationError += [String]$SPuser.LoginName
	}
	Finally
	{
		Remove-variable ADUserBySAMAccountName -ErrorAction SilentlyContinue
		Remove-variable ADUserBySID  -ErrorAction SilentlyContinue	
	}
}
#EndRegion

#Region Check AD account modification and launch update
Function CheckADAccountModificationAndUpdate
{
	Param
	(
		$_ADUserBySAMAccountName,
		$_ADUserBySID,
		$_SPuser,
		$_SPUserStr,
		$_SPUserSID,
		$_DomainName
	)

	Try
	{	
		If(($_ADUserBySAMAccountName -ne $null) -and ($_ADUserBySID -eq $null))
		{
			$ADUserBySAMAccountNameSID = $_ADUserBySAMAccountName.SID 
			Write-Host "  |--> Found $SPUserSAMAccountName account with different SID ($_SPUserSID <> $ADUserBySAMAccountNameSID)" -fore Red
			Write-Host "  |--> Update SharePoint User with new SID"
			$OldSPuser = $_SPuser
			
			UpdateUser -_Identity $_SPuser -_NewAlias $_SPUserStr

			If ($ADUserBySAMAccountNameSID -eq $_SPUserSID)
			{
				Write-Host "  |--> SharePoint user have been successfully updated." -fore Green
				$Global:ExecuteSynchronize = $True
				$Global:CounterUsersADAccountUpdateSuccess++
				$Global:ADUser = $_ADUserBySAMAccountName
			}
			Else
			{
				Write-Host "  |--> Failed to update SharePoint user. Synchronization of user have been aborted." -fore Red
				$Global:ExecuteSynchronize = $False
				$Global:CounterUsersADAccountUpdateFailed++
				$Global:UsersWithADAccountUpdateError += [String]$OldSPuser.LoginName
			}
			Remove-variable OldSPuser -ErrorAction SilentlyContinue
			Remove-variable ADUserBySAMAccountNameSID -ErrorAction SilentlyContinue
		}
		ElseIf(($_ADUserBySAMAccountName -eq $null) -and ($_ADUserBySID -ne $null))
		{
			$ADUserBySIDSAMAccountName = $_ADUserBySID.SAMAccountName
			
			Write-Host "  |--> Found $SPUserSAMAccountName account with different SAM Account Name ($SPUserSAMAccountName <> $ADUserBySIDSAMAccountName)"
			Write-Host "  |--> Update SharePoint User with new SAM Account Name"
			
			$UserNewLoginName = $_DomainName + "\" + $ADUserBySIDSAMAccountName
			$OldSPuser = $_SPuser								
			
			UpdateUser -_Identity $_SPuser -_NewAlias $UserNewLoginName
			
			If($_SPuser -ne $null)
			{
				Write-Host "  |--> SharePoint user have been successfully updated." -fore Green									
				$Global:ExecuteSynchronize = $True
				$Global:CounterUsersADAccountUpdateSuccess++
				$Global:ADUser = $_ADUserBySAMAccountName
			
			}
			Else
			{
				Write-Host "  |--> Failed to update sharePoint user. Synchronization of user have been aborted." -fore Red
				$Global:ExecuteSynchronize = $False
				$Global:CounterUsersADAccountUpdateFailed++
				$Global:UsersWithADAccountUpdateError += [String]$OldSPuser.LoginName
			}
			Remove-variable OldSPuser -ErrorAction SilentlyContinue
			Remove-variable ADUserBySAMAccountNameSID -ErrorAction SilentlyContinue
			Remove-variable ADUserBySIDSAMAccountName -ErrorAction SilentlyContinue
			Remove-variable UserNewLoginName -ErrorAction SilentlyContinue
		}
		Else
		{
			$ADUserBySIDSID = $_ADUserBySID.SID
			$ADUserBySAMAccountNameSID = $_ADUserBySAMAccountName.SID
			If($ADUserBySIDSID -eq $ADUserBySAMAccountNameSID)
			{
				Write-Host "  |--> $SPUserSAMAccountName account have been found" -fore Green
				$Global:ADUser = $_ADUserBySID
				$Global:ExecuteSynchronize = $True
				$Global:CounterUsersADAccountUpdateNoModification++										
			}
			Else
			{
				$ADUserBySIDSAMAccountName = $_ADUserBySID.SAMAccountName
				$ADUserBySAMAccountNameSAMAccountName = $_ADUserBySAMAccountName.SAMAccountName										
				Write-Host "  |--> Two account have been found with different SID" -fore Red
				Write-Host "  |--> Account found by SID : $ADUserBySIDSID / $ADUserBySIDSAMAccountName" -fore Red
				Write-Host "  |--> Account found by SAM Account Name : $ADUserBySAMAccountNameSID / $ADUserBySAMAccountNameSAMAccountName" -fore Red
				Write-Host "  |--> Synchronization of user have been aborted." -fore Red
				$Global:ExecuteSynchronize = $False
				$Global:CounterUsersADAccountUpdateFailed++
				$Global:UsersWithADAccountUpdateError += [String]$_SPuser.LoginName
				Remove-variable ADUserBySIDSAMAccountName -ErrorAction SilentlyContinue
				Remove-variable ADUserBySAMAccountNameSAMAccountName -ErrorAction SilentlyContinue											
			}
			Remove-variable ADUserBySIDSID -ErrorAction SilentlyContinue
			Remove-variable ADUserBySAMAccountNameSID -ErrorAction SilentlyContinue									
		}											
	}
	Catch
	{
		$errText = $error[0].Exception.Message
		Write-Host "  |--> Failed to check AD account modification. Reason: $errText" -fore Red	
	}
}
#EndRegion

#Region Update SharePoint User
Function UpdateUser
{
	Param
	(
		$_Identity,
		$_NewAlias
	)

	Try
	{
		Move-SPUser -Identity $_Identity -newalias $_NewAlias -IgnoreSID -Confirm:$false -ErrorAction Stop
		Retrieve-User-And-Properties $_Identity
	}
	Catch
	{
		$Global:SPuser = $null
		$errText = $error[0].Exception.Message
		Write-Host "  |--> Failed to update SharePoint User. Reason: $errText" -fore Red
	}
}
#EndRegion

#Region Control if user attributes have been modified"
Function NativeSynchronization
{
	Param
	(
		$_Identity,
		$_Web
	)
	Try
	{
		Write-Host "  |--> Get current user attributes"
		
		GetCurrentUserAttributes -_Identity $_Identity -_web $_Web
		
		Write-Host "  |--> Synchronize with Set-SPuser and SyncFromAD parameter"
		
		Set-SPUser -Identity $_Identity -web $_Web -SyncFromAD -ErrorAction Stop

		Write-Host "  |--> Control if user attributes have been modified"
		ControlUserAttributesModification -_Identity $_Identity -_web $_Web
	}
	Catch
	{
		$errText = $error[0].Exception.Message
		Write-Host "  |--> User synchronization has failed. Reason: $errText" -fore Red
		$Global:CounterUsersNativeSynchronizationFailed++
		$Global:UsersWithNativeSynchonizationError += [String]$_Identity.LoginName
	}
	Finally
	{
		Remove-variable NewUserInfo -ErrorAction SilentlyContinue
		Remove-variable NewUserLogin -ErrorAction SilentlyContinue
		Remove-variable NewUserdisplayName -ErrorAction SilentlyContinue
		Remove-variable NewUserName -ErrorAction SilentlyContinue
		Remove-variable NewUserEmail -ErrorAction SilentlyContinue
		Remove-variable NewUserLoginName -ErrorAction SilentlyContinue
						
		Remove-variable OldUserInfo -ErrorAction SilentlyContinue
		Remove-variable OldUserLogin -ErrorAction SilentlyContinue
		Remove-variable OldUserdisplayName -ErrorAction SilentlyContinue
		Remove-variable OldUserName -ErrorAction SilentlyContinue
		Remove-variable OldUserEmail -ErrorAction SilentlyContinue
		Remove-variable OldUserLoginName -ErrorAction SilentlyContinue	
	}
}
#EndRegion

#Region Get current user attributes
Function GetcurrentUserAttributes
{
	Param
	(
		$_Identity,
		$_Web
	)
	Try
	{
		$Global:OldUserInfo = Get-SPUser -Identity $_Identity -web $_Web
		$Global:OldUserLogin = $OldUserInfo.UserLogin
		$Global:OldUserdisplayName = $OldUserInfo.DisplayName
		$Global:OldUserName = $OldUserInfo.Name
		$Global:OldUserEmail = $OldUserInfo.Email
		$Global:OldUserLoginName = $OldUserInfo.LoginName
	}
	Catch
	{
		$errText = $error[0].Exception.Message
		Write-Host "  |--> Failed to get current user attributes. Reason: $errText" -fore Red	
	}
}
#EndRegion

#Region Control if user attributes have been modified"
Function ControlUserAttributesModification
{
	Param
	(
		$_Identity,
		$_Web
	)
	Try
	{
		$Global:NewUserInfo = Get-SPUser -Identity $_Identity -web $_Web
		$Global:NewUserLogin = $NewUserInfo.UserLogin
		$Global:NewUserdisplayName = $NewUserInfo.DisplayName
		$Global:NewUserName = $NewUserInfo.Name
		$Global:NewUserEmail = $NewUserInfo.Email
		$Global:NewUserLoginName = $NewUserInfo.LoginName
		$Global:UserModified = $False
		
		If ($DebugMode -eq $True)
		{
			Write-host "----Debug => SPuser"
			Write-host "SPuser all properties:" $NewUserInfo | select *
			Write-host "----Debug => OldValue"
			Write-host "OldUserLogin:" $OldUserLogin
			Write-host "OldUserdisplayName:"$OldUserdisplayName
			Write-host "OldUserName:"$OldUserName
			Write-host "OldUserEmail:"$OldUserEmail
			Write-host "OldUserLoginName:"$OldUserLoginName
			Write-host "----Debug => NewValue"
			Write-host "NewUserLogin:"$NewUserLogin
			Write-host "NewUserdisplayName:"$NewUserdisplayName
			Write-host "NewUserName:"$NewUserName
			Write-host "NewUserEmail:"$NewUserEmail
			Write-host "NewUserLoginName:"$NewUserLoginName
		}
		
		If ($OldUserLogin -ne $NewUserLogin)
		{
			Write-Host "  |--> User Login has been modified ($OldUserLogin ==> $NewUserLogin)" -fore Green
			$UserModified = $True
		}
		If ($OldUserdisplayName -ne $NewUserdisplayName)
		{
			Write-Host "  |--> User Display Name has been modified ($OldUserdisplayName ==> $NewUserdisplayName)" -fore Green
			$UserModified = $True
		}
		If ($OldUserName -ne $NewUserName)
		{
			Write-Host "  |--> User Name has been modified ($OldUserName ==> $NewUserName)" -fore Green
			$UserModified = $True
		}
		If ($OldUserEmail -ne $NewUserEmail)
		{
			Write-Host "  |--> User Email has been modified ($OldUserEmail ==> $NewUserEmail)" -fore Green
			$UserModified = $True
		}					
		If ($OldUserLoginName -ne $NewUserLoginName)
		{
			Write-Host "  |--> User Login Name has been modified ($OldUserLoginName ==> $NewUserLoginName)" -fore Green
			$UserModified = $True
		}
		If ($UserModified -eq $False)
		{
			Write-Host "  |--> User hasn't been modified" -fore Green
			$Global:CounterUsersNativeSynchronizationNoModification++
		}
		Else
		{
			$Global:CounterUsersNativeSynchronizationSuccess++
		}
	}
	Catch
	{
		$errText = $error[0].Exception.Message
		Write-Host "  |--> Failed to control current user attributes modification. Reason: $errText" -fore Red	
	}
}
#EndRegion

#Region Get user extended information from SharePoint
Function GetUserExtendedInformation
{
	Param
	(
		$_SPUserStr,
		$_Web
	)
	Try
	{
		$Global:list = $_Web.Lists["User Information List"]
		$Global:query = New-Object Microsoft.SharePoint.SPQuery
		$Global:query.Query = "<Where><Eq><FieldRef Name='Name' /><Value Type='Text'>$_SPUserStr</Value></Eq></Where>"
	}
	Catch
	{
		$errText = $error[0].Exception.Message
		$Global:List = $Null
		Write-Host "  |--> Failed to get user extended information from SharePoint. Reason: $errText" -fore Red	
	}
}
#EndRegion

#Region Check user extended information from SharePoint
Function CheckUserExtendedInformation
{
	Param
	(
		$_list,
		$_query
	)

	foreach ($item in $_list.GetItems($_query)) 
	{
		Try
		{	
			$Global:OldUserJobTitle = $item["JobTitle"]
			$Global:OldUserDepartment = $item["Department"]
			$Global:OldUserIPPhone = $item["IPPhone"]							
			$Global:OldUserMobilePhone = $item["MobilePhone"]
			$Global:OldUserTitle = $item["Title"]
			Remove-variable item -ErrorAction SilentlyContinue	
		}
		Catch
		{
			$errText = $error[0].Exception.Message
			Write-Host "  |--> Failed to check user extended information from SharePoint. Reason: $errText" -fore Red	
		}	
	}

}
#EndRegion

#Region Update user extended information from SharePoint
Function UpdateUserExtendedInformation
{
	Param
	(
		$_list,
		$_query,
		$_ADUser
	)

	foreach ($item in $_list.GetItems($_query)) 
	{
		Try
		{	
			If ($DebugMode -eq $True)
			{						
				Write-Host "SP User Jobtitle :"$item["JobTitle"]
				[string]::IsNullOrEmpty($item["JobTitle"])
				Write-Host "AD User Title :"$_ADUser.title
				[string]::IsNullOrEmpty($_ADUser.title)
			}
			
			If((![string]::IsNullOrEmpty($item["JobTitle"])) -and (![string]::IsNullOrEmpty($_ADUser.title)) -and ($item["JobTitle"] -ne [string]$_ADUser.title))
			{
				Write-Host "  |--> Job Title must be updated"
				$item["JobTitle"] = [string]$_ADUser.title
			}

			If ($DebugMode -eq $True)
			{						
				Write-Host "SP User Department :"$item["Department"]							
				Write-Host "AD User Department :"$_ADUser.department
			}							
			
			If((![string]::IsNullOrEmpty($item["Department"])) -and (![string]::IsNullOrEmpty($_ADUser.department)) -and ($item["Department"] -ne [string]$_ADUser.department))
			{
				Write-Host "  |--> Department must be updated"
				$item["Department"] = [string]$_ADUser.department
			}
			
			If ($DebugMode -eq $True)
			{						
				Write-Host "SP User IPPhone:"$item["IPPhone"]							
				Write-Host "AD User OfficePhone :"$_ADUser.OfficePhone
			}
			
			If((![string]::IsNullOrEmpty($item["OfficePhone"])) -and (![string]::IsNullOrEmpty($_ADUser.OfficePhone)) -and ($item["IPPhone"] -ne [string]$_ADUser.OfficePhone))
			{
				Write-Host "  |--> Office Phone must be updated"
				$item["IPPhone"] = [string]$_ADUser.OfficePhone
			}
			
			If ($DebugMode -eq $True)
			{						
				Write-Host "SP User MobilePhone:"$item["MobilePhone"]							
				Write-Host "AD User MobilePhone :"$_ADUser.MobilePhone
			}
			
			If((![string]::IsNullOrEmpty($item["MobilePhone"])) -and (![string]::IsNullOrEmpty($_ADUser.MobilePhone)) -and ($item["MobilePhone"] -ne [string]$_ADUser.MobilePhone))
			{
				Write-Host "  |--> Mobile Phone must be updated"							
				$item["MobilePhone"] = [string]$_ADUser.mobile
			}

			If ($DebugMode -eq $True)
			{						
				Write-Host "SP User Title :"$item["Title"]							
				Write-Host "AD User DisplayName :"$_ADUser.DisplayName
			}
			
			If((![string]::IsNullOrEmpty($item["Title"])) -and (![string]::IsNullOrEmpty($_ADUser.DisplayName)) -and ($item["Title"] -ne [string]$_ADUser.DisplayName))
			{
				Write-Host "  |--> Title must be updated"
				$item["Title"]= [string]$_ADUser.DisplayName
			}
			
			$item.SystemUpdate()
			Remove-variable item -ErrorAction SilentlyContinue	
		}
		Catch
		{
			$errText = $error[0].Exception.Message
			Write-Host "  |--> Failed to update user extended information from SharePoint. Reason: $errText" -fore Red	
		}	
	}

}
#EndRegion

#Region Check user extended information modification
Function CheckUserExtendedInformationModification
{
	Param
	(
		$_list,
		$_query,
		$_ADUser
	)

	foreach ($item in $_list.GetItems($_query)) 
	{
		Try
		{	
			If((![string]::IsNullOrEmpty($item["JobTitle"])) -and (![string]::IsNullOrEmpty($_ADUser.title)) -and ($item["JobTitle"] -ne [string]$_ADUser.title))
			{
				$ADUserTitle = [string]$_ADUser.title
				Write-Host "  |--> Failed to modify user Job Title ($OldUserJobTitle <> $ADUserTitle)" -fore Red
				$Global:UserAdvancedModificationFailed = $True
				Remove-variable ADUserTitle -ErrorAction SilentlyContinue
			}
			Else
			{
				If($OldUserJobTitle -ne $item["JobTitle"])
				{
					$NewUserJobTitle = $item["JobTitle"]
					Write-Host "  |--> User Job Title has been modified ($OldUserJobTitle ==> $NewUserJobTitle)" -fore Green
					$UserAdvancedModified = $True
					Remove-variable NewUserJobTitle -ErrorAction SilentlyContinue
				}
			}
			Remove-variable OldUserJobTitle -ErrorAction SilentlyContinue
			
			If((![string]::IsNullOrEmpty($item["Department"])) -and (![string]::IsNullOrEmpty($_ADUser.department)) -and ($item["Department"] -ne [string]$_ADUser.department))
			{
				$ADUserDepartment = [string]$_ADUser.department
				Write-Host "  |--> Failed to modify user Department ($OldUserDepartment <> $ADUserDepartment)" -fore Red
				$Global:UserAdvancedModificationFailed = $True
				Remove-variable ADUserDepartment -ErrorAction SilentlyContinue									
			}
			Else
			{
				If($OldUserDepartment -ne $item["Department"])
				{
					$NewUserDepartment = $item["Department"]
					Write-Host "  |--> User Department has been modified ($OldUserDepartment ==> $NewUserDepartment)" -fore Green
					$UserAdvancedModified = $True
					Remove-variable NewUserDepartment -ErrorAction SilentlyContinue
				}
			}
			Remove-variable OldUserDepartment -ErrorAction SilentlyContinue
											
			If((![string]::IsNullOrEmpty($item["IPPhone"])) -and (![string]::IsNullOrEmpty($_ADUser.OfficePhone)) -and ($item["IPPhone"] -ne [string]$_ADUser.OfficePhone))
			{
				$ADUserOfficePhone = [string]$_ADUser.OfficePhone
				Write-Host "  |--> Failed to modify user IP Phone ($OldUserIPPhone <> $ADUserOfficePhone)" -fore Red
				$Global:UserAdvancedModificationFailed = $True
				Remove-variable ADUserOfficePhone -ErrorAction SilentlyContinue
			}
			Else
			{
				If($OldUserIPPhone -ne $item["IPPhone"])
				{
					$NewUserIPPhone = $item["IPPhone"]
					Write-Host "  |--> User IP Phone has been modified ($OldUserIPPhone ==> $NewUserIPPhone)" -fore Green
					$UserAdvancedModified = $True
					Remove-variable NewUserIPPhone -ErrorAction SilentlyContinue
				}
			}
			Remove-variable OldUserIPPhone -ErrorAction SilentlyContinue
											
			If((![string]::IsNullOrEmpty($item["MobilePhone"])) -and (![string]::IsNullOrEmpty($_ADUser.MobilePhone)) -and ($item["MobilePhone"] -ne [string]$_ADUser.MobilePhone))
			{
				$ADUserMobilePhone = [string]$_ADUser.MobilePhone
				Write-Host "  |--> Failed to modify user Mobile Phone ($OldUserMobilePhone <> $ADUserMobilePhone)" -fore Red
				$Global:UserAdvancedModificationFailed = $True
				Remove-variable ADUserMobilePhone -ErrorAction SilentlyContinue
			}
			Else
			{
				If($OldUserMobilePhone -ne $item["MobilePhone"])
				{
					$NewUserMobilePhone = $item["MobilePhone"]
					Write-Host "  |--> User Mobile Phone has been modified ($OldUserMobilePhone ==> $NewUserMobilePhone)" -fore Green
					$UserAdvancedModified = $True
					Remove-variable NewUserMobilePhone -ErrorAction SilentlyContinue
				}
			}
			Remove-variable OldUserMobilePhone -ErrorAction SilentlyContinue
											
			If((![string]::IsNullOrEmpty($item["Title"])) -and (![string]::IsNullOrEmpty($_ADUser.DisplayName)) -and ($item["Title"] -ne [string]$_ADUser.DisplayName))
			{
				$ADUserDisplayName = [string]$_ADUser.DisplayName
				Write-Host "  |--> Failed to modify user Title ($OldUserTitle <> $ADUserDisplayName)" -fore Red
				$Global:UserAdvancedModificationFailed = $True
				Remove-variable ADUserDisplayName -ErrorAction SilentlyContinue
			}
			Else
			{
				If($OldUserTitle -ne $item["Title"])
				{
					$NewUserTitle = $item["Title"]
					Write-Host "  |--> User Title has been modified ($OldUserTitle ==> $NewUserTitle)" -fore Green
					$UserAdvancedModified = $True
					Remove-variable NewUserTitle -ErrorAction SilentlyContinue
				}
			}
			Remove-variable OldUserTitle -ErrorAction SilentlyContinue
										
			If ($UserAdvancedModified -eq $False)
			{
				If($UserAdvancedModificationFailed -eq $False)
				{
					Write-Host "  |--> SharePoint user extended information hasn't been modified" -fore Green
					$Global:CounterUsersAdvancedSynchronizationNoModification++
				}
				Else
				{
					Write-Host "  |--> Failed to modify SharePoint user extended information" -fore Red
					$Global:CounterUsersAdvancedSynchronizationFailed++									
				}
			}
			Else
			{
				$Global:CounterUsersAdvancedSynchronizationSuccess++
			}
			Remove-variable item -ErrorAction SilentlyContinue												
		}
		Catch
		{
			$errText = $error[0].Exception.Message
			Write-Host "  |--> Failed to check user extended information modification. Reason: $errText" -fore Red	
		}	
	}

}
#EndRegion

#Region Send e-mail to nominated recipient(s)
function Send-Email 
{			
	try
	{
		$msg = new-object Net.Mail.MailMessage
		$msg.From = $emailFrom 
		
		if ($emailTo -ne "") 
		{
			$msg.To.Add($emailTo)
		}
		else
		{
			Write-Host "|--> E-mail was not sent. Reason: There is no nominated recipient." -fore Red
			break	
		}

		if ($emailCC -ne "") 
		{
			$msg.CC.Add($emailCC)
		}
		
		If ($jobStatus.contains("error") -eq "True")
			{
				if ($emailOnErrorTO -ne "") {$msg.To.Add($emailOnErrorTO)}
				if ($emailOnErrorCC -ne "") {$msg.To.Add($emailOnErrorCC)}
			}
			
		$msg.Subject = $JobStatus
		$msg.Body = $CompleteJobSummary
		$msg.IsBodyHtml = $true
		
		#Attach file, if applicable
		if ($fileAttachment -ne $null)
		{
			$fileAttachment = New-Object Net.Mail.Attachment($fileAttachment)
			$msg.Attachments.Add($fileAttachment)
		}

		$smtp = new-object Net.Mail.SmtpClient($smtpServer)
		$smtp.Send($msg)
		Write-Host "|--> E-mail was sent." -fore Green
	}
	catch 
	{
		$errText = $error[0].Exception.Message
		Write-Host "|--> Sending mail to one or more recipients failed. Reason: $errText" -fore Red
	}	
}
#EndRegion

#*===============================================================================
# Main
#*===============================================================================

# Check if the execution policy is set to Unrestricted   
try
{	$policy = Get-ExecutionPolicy   

	if($policy -ne "Unrestricted")
	{   
		Set-ExecutionPolicy "RemoteSigned" 
	} 
}
catch
{
		$errText = $error[0].Exception.Message
		Write-Host "|--> A problem occurred whilst attempting to set the the Execution Policy. Reason: $errText" -fore Red
}
	
$StartTime = (Get-Date)

Test-FilePath-Create $Log

Start-Transcript $Log
if (!$?)
{
	Write-Host "Transcript cannot start because path is unavailable" -Fore Red
	$TranscriptStatus= "Error"
}

Write-host "----------------------------------------------------------------------------" -fore Yellow
Write-host "--           SharePoint Foundation User Synchronization                   --" -fore Yellow
Write-host "----------------------------------------------------------------------------" -fore Yellow
Write-host "|--> Load sharePoint snap-in"

Load-Snapin

Write-Host "|--> Import Active Directory Module"

LoadActiveDirectoryModule

Write-host "|--> List sites"

$sites = Get-SPSite -Limit ALL

Write-host "|--> Sites to process :"

$sites

$Global:DomainReachable = @()
$Global:DomainUnReachable = @()
$Global:GlobalResult = @{}

foreach($site in $sites) {

    $web = $site.RootWeb
    if($web -ne $null) {
	
		$SiteUrl =  $site.url
		Write-host "============================================================================" -fore Yellow
		Write-host "== Start user synchronization for $SiteUrl ($web)" -fore Yellow
		Write-host "============================================================================" -fore Yellow
		Write-Host "|--> List users"
		
		$CounterUsersToSynchronize = 0
		$Global:CounterUsersDomainUnreachable = 0
		$Global:CounterUsersAdvancedNotFound = 0		
		$Global:CounterUsersADAccountUpdateSuccess = 0
		$Global:CounterUsersADAccountUpdateFailed = 0
		$Global:CounterUsersADAccountUpdateNoModification = 0
		$Global:CounterUsersNativeSynchronizationSuccess = 0
		$Global:CounterUsersNativeSynchronizationFailed = 0
		$Global:CounterUsersNativeSynchronizationNoModification = 0
		$Global:CounterUsersAdvancedSynchronizationSuccess = 0
		$Global:CounterUsersAdvancedSynchronizationFailed = 0
		$Global:CounterUsersAdvancedSynchronizationNoModification = 0
		$CounterUsersDeletionSuccess = 0
		$CounterUsersDeletionFailed = 0
		$Global:UsersWithDomainUnreachable = @()
		$Global:UsersNotFound = @()
		$Global:UsersWithADAccountUpdateError = @()
		$Global:UsersWithNativeSynchonizationError = @()
		$Global:UsersWithAdvancedSynchonizationError  = @()
		
		#Regex to exclude system account and SharePoint group
		$RegExclusionList = "c:0\(\.s\|truec:0\(\.s\|true|c:0!\.s\|windows|c:0!\.s\|forms:aspnetsqlmembershipprovider|NT\ AUTHORITY\\.*|SHAREPOINT\\.*"
		
        $Users = Get-SPUser -Limit All -web $web | where-object {$_.IsDomainGroup -eq $False}| Where-object{ $_.LoginName -notmatch $RegExclusionList}
		
		$UsersToSynchronize = $Users.count
		
		Write-Host "|--> Users to synchronize : "$UsersToSynchronize
		
		foreach ($User in $Users)
		{
			Retrieve-User-And-Properties $User
            if ($SPuser -ne $null)
            {
				If ($DebugMode -eq $True)
				{						
					Write-Host "SP User properties) :" ($SPuser | select *)
					Write-Host "SP User SAM Account Name :"$SPUserSAMAccountName
					Write-Host "SP User Domain Name :"$DomainName
					Write-Host "SP User SID :"$SPUserSID					
				}				
				
				$CounterUsersToSynchronize++
				
				Write-Host "|--> Synchronize $SPuser ($CounterUsersToSynchronize/$UsersToSynchronize)"

				Write-host "  |--> Get list of DCs in domain '$DomainName'"
				
				# Test if user's domain is reachable, one time by domain for all users to synchronize
				TestDomainAvailability -_DomainName $DomainName
				
				If($DomainTestedWithSuccess -eq $True)
				{
					If($ImportModuleAD -eq $False)
					{				
						#Launch native synchronisation with control of modification
						NativeSynchronization -_Identity $SPuser -_web $web
					}
					Else
					{
						Write-Host "  |--> Get user information from domain"

						GetAndCheckADAccountModification -_SPUserSAMAccountName $SPUserSAMAccountName -_SPUserSID $SPUserSID -_SPuser $SPuser -_SPUserStr $SPUserStr -_DomainName $DomainName

						If($ExecuteSynchronize -eq $True)
						{
							#Launch native synchronisation with control of modification
							NativeSynchronization -_Identity $SPuser -_web $web
							
							Write-Host "  |--> Get user extended information from SharePoint"
							
							GetUserExtendedInformation -_SPUserStr $SPUserStr -_Web $web
							
							Write-Host "  |--> Synchronize extended information"
							If ($List -ne $null)
							{
								CheckUserExtendedInformation -_list $list -_query $query
								
								UpdateUserExtendedInformation -_list $list -_query $query -_ADUser $ADUser
								
								Remove-variable List -ErrorAction SilentlyContinue
								Remove-variable query -ErrorAction SilentlyContinue
								
								Write-Host "  |--> Control if user extended information have been modified"
								
								$Global:UserAdvancedModified = $False
								$Global:UserAdvancedModificationFailed = $False
								
								GetUserExtendedInformation -_SPUserStr $SPUserStr -_Web $web
								
								CheckUserExtendedInformationModification -_list $list -_query $query -_ADUser $ADUser
							
								Remove-variable List -ErrorAction SilentlyContinue
								Remove-variable query -ErrorAction SilentlyContinue
							}
							Else
							{
								Write-Host "  |--> Failed to get $SPUserStr SharePoint user information" -fore Red
								$Global:CounterUsersAdvancedSynchronizationFailed++
								$Global:UsersWithAdvancedSynchonizationError += [String]$SPuser.LoginName
							}
							Remove-variable ADUser -ErrorAction SilentlyContinue
						}							
					}
				}
				Remove-variable DomainTestedWithSuccess -ErrorAction SilentlyContinue	
				Remove-variable SPUserStr -ErrorAction SilentlyContinue
				Remove-variable SplitSPuser -ErrorAction SilentlyContinue
				Remove-variable SPUserSAMAccountName -ErrorAction SilentlyContinue
				Remove-variable SPUserSID -ErrorAction SilentlyContinue
				Remove-variable DomainName -ErrorAction SilentlyContinue					
				Remove-variable Nltestexe -ErrorAction SilentlyContinue
				Remove-variable NltestParam -ErrorAction SilentlyContinue
				Remove-variable NlTestResult -ErrorAction SilentlyContinue				
            }
			Remove-variable SPuser -ErrorAction SilentlyContinue
			Remove-variable User -ErrorAction SilentlyContinue
			[GC]::Collect()			
        }
    }
	
    $web.Dispose()
    $site.Dispose()
	Write-host "============================================================================" -fore Yellow
	Write-host "== Synchronization Report for $SiteUrl ($web)" -fore Yellow
	Write-host "============================================================================" -fore Yellow
	Write-host "| Users to synchronize : "$UsersToSynchronize
	Write-host "----------------------------------------------------------------------------"
	Write-host "| Users with unreachable domain : "	$CounterUsersDomainUnreachable
	If ($CounterUsersDomainUnreachable -ne 0)
	{
		Write-host "----------------------------------------------------------------------------"
		Write-host "| List of user for which the domain is unreachable :"
		Write-Host "| "
		$UsersWithDomainUnreachable | %{Write-host "| "$_}
	}
	Write-host "----------------------------------------------------------------------------"
	
	If($ImportModuleAD -eq $False)
	{	
		Write-host "| Users Native synchronization Succes : "$CounterUsersNativeSynchronizationSuccess
		Write-host "| Users Native synchronization Failed : "$CounterUsersNativeSynchronizationFailed
		Write-host "| Users Native synchronization with no modification : "$CounterUsersNativeSynchronizationNoModification
		If($CounterUsersNativeSynchronizationFailed -ne 0)
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of user for which the native synchronization failed :"
			Write-Host "| "
			$UsersWithNativeSynchonizationError | %{Write-host "| "$_}
		}
		$GlobalResult.add($SiteUrl,[math]::Round((($UsersToSynchronize-($CounterUsersNativeSynchronizationFailed+$CounterUsersDomainUnreachable))/$UsersToSynchronize)*100,1))
	}
	Else
	{
		Write-host "| Users not found : "$CounterUsersAdvancedNotFound
		Write-host "| Users AD account update Succes : "$CounterUsersADAccountUpdateSuccess
		Write-host "| Users AD account update Failed : "$CounterUsersADAccountUpdateFailed
		Write-host "| Users AD account with no modification : "$CounterUsersADAccountUpdateNoModification			
		Write-host "| Users Native synchronization Succes : "$CounterUsersNativeSynchronizationSuccess
		Write-host "| Users Native synchronization Failed : "$CounterUsersNativeSynchronizationFailed
		Write-host "| Users Native synchronization with no modification : "$CounterUsersNativeSynchronizationNoModification		
		Write-host "| Users Advanced synchronization Succes : "$CounterUsersAdvancedSynchronizationSuccess
		Write-host "| Users Advanced synchronization Failed : "$CounterUsersAdvancedSynchronizationFailed
		Write-host "| Users Advanced synchronization with no modification : "$CounterUsersAdvancedSynchronizationNoModification
		If($CounterUsersAdvancedNotFound -ne 0)
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of users who have not been found in domain :"
			Write-Host "| "
			$UsersNotFound | %{Write-host "| "$_}
		}
		If($CounterUsersADAccountUpdateFailed -ne 0)
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of users for which AD account update failed :"
			Write-Host "| "
			$UsersWithADAccountUpdateError | %{Write-host "| "$_}
		}		
		If($CounterUsersNativeSynchronizationFailed -ne 0)
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of user for which the native synchronization failed :"
			Write-Host "| "
			$UsersWithNativeSynchonizationError | %{Write-host "| "$_}		
		}
		If($CounterUsersAdvancedSynchronizationFailed -ne 0 )
		{
			Write-host "----------------------------------------------------------------------------"
			Write-host "| List of user for which the advanced synchronization failed :"
			Write-Host "| "
			$UsersWithAdvancedSynchonizationError | %{Write-host "| "$_}
		}
		$GlobalResult.add([String]$SiteUrl.replace("SPSite Url=",""),[math]::Round((($UsersToSynchronize-($CounterUsersNativeSynchronizationFailed+$CounterUsersAdvancedSynchronizationFailed+$CounterUsersDomainUnreachable+$CounterUsersAdvancedNotFound+$CounterUsersADAccountUpdateFailed))/$UsersToSynchronize)*100,1))
	}

	If(($DeleteUSersNotFound -eq $True) -and (($UsersNotFound.count -gt 0) -or ($UsersWithDomainUnreachable.count -gt 0)))
	{
		Write-host "============================================================================" -fore Yellow
		Write-host "== Deleting user that are not found or with unreachable domain for $SiteUrl ($web)" -fore Yellow
		Write-host "============================================================================" -fore Yellow	
		$DeletedUsersRatio = [System.Math]::Round((($CounterUsersAdvancedNotFound+$CounterUsersDomainUnreachable)/$UsersToSynchronize)*100,1)
		Write-host "|-> $DeletedUsersRatio % of users have not found"
		#To avoid a mass removal in case of unavailability of a domain, I introduced a maximum ratio of account to be deleted
		If($DeletedUsersRatio -le 30)
		{
			Write-host "|-> The ratio is below than 30% so we will proceed with the removal of accounts" -fore Green
			$UsersNotFound += $UsersWithDomainUnreachable
			foreach ($UserToDelete in $UsersNotFound)
			{
				Write-host "|-> Remove of user"$UserToDelete
				
				Try
				{
					Remove-SPUser -identity $UserToDelete -web $SiteUrl -Confirm:$False  -ErrorAction Stop
					$TestUser = Get-SPuser -Limit ALL -web $SiteUrl | where-object {$_.IsDomainGroup -eq $False} | Where-object {$_ -eq $UserToDelete}
					If($TestUser -ne $null)
					{
						Write-host "|-> Failed to remove "$UserToDelete -fore Red
						$CounterUsersDeletionFailed++
					}
					Else
					{
						Write-Host $UserToDelete "has been removed" -fore Green
						$CounterUsersDeletionSuccess++
					}
				}
				Catch
				{
					$errText = $error[0].Exception.Message
					Write-host "|-> Failed to remove "$UserToDelete". Reason: $errText" -fore Red
					$CounterUsersDeletionFailed++
				}
			}
		}
		Else
		{
			Write-host "|-> The ratio is higher than 30% so we won't proceed with the removal of accounts" -fore Red
		}
	}
	If($SendMail -eq $True)
	{
		#Creating an html report for this site collection
		$jobSummary +="		
		<p><strong><u>Synchronization Report for $SiteUrl ($web):</u></strong></p>
		<table border='1' cellpadding='0' cellspacing='0'>
				<tbody>
					<tr>
						<td>
						<p>Users to synchronize</p>
						</td>
						<td>
						<p>$UsersToSynchronize</p>
						</td>
					</tr>
					<tr>
						<td>
						<p>Users with unreachable domain</p>
						</td>
						<td>
						<p>$CounterUsersDomainUnreachable</p>
						</td>
					</tr>
					<tr>
						<td>
						<p>List of user for which the domain is unreachable</p>
						</td>
						<td>
		"
		If ($CounterUsersDomainUnreachable -ne 0)
		{
			$UsersWithDomainUnreachable | %{$jobSummary += $_ + "<br />"}
		}
		
		$jobSummary +="
						</td>
					</tr>				
		"
		If($ImportModuleAD -eq $False)
		{
			$jobSummary +="
					<tr>
						<td>
						<p>Users Native synchronization Succes</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationSuccess</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization Failed</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationFailed</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization with no modification</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationNoModification</p>						
					</tr>
					<tr>
						<td>
						<p>List of user for which the native synchronization failed</p>
						</td>
						<td>					
			"
			If($CounterUsersNativeSynchronizationFailed -ne 0)
			{
				$UsersWithNativeSynchonizationError | %{$jobSummary += $_ + "<br />"}
			}
			$jobSummary +="
							</td>
						</tr>
					</tbody>
				</table>						
			"
		}
		Else
		{
			$jobSummary +="
					<tr>
						<td>
						<p>Users not founds</p>
						</td>
						<td>
						<p>$CounterUsersAdvancedNotFound</p>						
					</tr>
					<tr>
						<td>
						<p>Users AD account update Success</p>
						</td>
						<td>
						<p>$CounterUsersADAccountUpdateSuccess</p>						
					</tr>
					<tr>
						<td>
						<p>Users AD account update Failed</p>
						</td>
						<td>
						<p>$CounterUsersADAccountUpdateFailed</p>						
					</tr>
					<tr>
						<td>
						<p>Users AD account with no modification</p>
						</td>
						<td>
						<p>$CounterUsersADAccountUpdateNoModification</p>						
					</tr>					
					<tr>
						<td>
						<p>Users Native synchronization Success</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationSuccess</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization Failed</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationFailed</p>						
					</tr>
					<tr>
						<td>
						<p>Users Native synchronization with no modification</p>
						</td>
						<td>
						<p>$CounterUsersNativeSynchronizationNoModification</p>						
					</tr>
					<tr>
						<td>
						<p>Users Advanced synchronization Success</p>
						</td>
						<td>
						<p>$CounterUsersAdvancedSynchronizationSuccess</p>						
					</tr>
					<tr>
						<td>
						<p>Users Advanced synchronization Failed</p>
						</td>
						<td>
						<p>$CounterUsersAdvancedSynchronizationFailed</p>						
					</tr>
					<tr>
						<td>
						<p>Users Advanced synchronization with no modification </p>
						</td>
						<td>
						<p>$CounterUsersAdvancedSynchronizationNoModification</p>						
					</tr>					
					<tr>
						<td>
						<p>List of users who have not been found in domain</p>
						</td>
						<td>				
			"		
			If($CounterUsersAdvancedNotFound -ne 0)
			{
				$UsersNotFound | %{$jobSummary += $_ + "<br />"}
			}
			$jobSummary +="
						</td>
					</tr>
					<tr>
						<td>
						<p>List of user for which AD account update failed</p>
						</td>
						<td>						
			"
			If($CounterUsersADAccountUpdateFailed -ne 0)
			{
				$UsersWithADAccountUpdateError | %{$jobSummary += $_ + "<br />"}		
			}			
			$jobSummary +="
						</td>
					</tr>
					<tr>
						<td>
						<p>List of user for which the native synchronization failed</p>
						</td>
						<td>						
			"
			If($CounterUsersNativeSynchronizationFailed -ne 0)
			{
				$UsersWithNativeSynchonizationError | %{$jobSummary += $_ + "<br />"}		
			}
			$jobSummary +="
						</td>
					</tr>
					<tr>
						<td>
						<p>List of user for which the advanced synchronization failed</p>
						</td>
						<td>						
			"			
			If($CounterUsersAdvancedSynchronizationFailed -ne 0 )
			{
				$UsersWithAdvancedSynchonizationError | %{$jobSummary += $_ + "<br />"}
			}
			$jobSummary +="
						</td>
					</tr>"
			If(($DeleteUSersNotFound -eq $True) -and (($UsersNotFound.count -gt 0) -or ($UsersWithDomainUnreachable.count -gt 0)))
			{
				$DeletedUsersRatio = [System.Math]::Round((($CounterUsersAdvancedNotFound+$CounterUsersDomainUnreachable)/$UsersToSynchronize)*100,1)
				$jobSummary +="<tr>
							<td>
							<p>$DeletedUsersRatio % of users have not found or have unreachable domain</p>
							</td>"				
				If($DeletedUsersRatio -le 30)
				{
					$jobSummary +="
								<td>
								<p><font color='green'>The ratio is less than 30% so we will proceed to the removal of accounts</font></p>
								</td>
							</tr>
							<tr>
								<td>
								<p>Number of users deleted with success</p>
								</td>
								<td>
								$CounterUsersDeletionSuccess
								</td>
							</tr>
							<tr>
								<td>
								<p>Number of users deleted with errors</p>
								</td>
								<td>
								$CounterUsersDeletionFailed
								</td>
							</tr>"
				}
				Else
				{
					$jobSummary +="
								<td>
								<p><font color='Red'>The ratio is more than 30% so we won't' proceed to the removal of accounts</font></p>
								</td>
							</tr>"					
				}
			}					
			$jobSummary +="
				</tbody>
			</table>
			"
		}		
	}
	Remove-variable UsersToSynchronize -ErrorAction SilentlyContinue
	Remove-variable	CounterUsersDomainUnreachable -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedNotFound -ErrorAction SilentlyContinue	
	Remove-variable CounterUsersADAccountUpdateSuccess -ErrorAction SilentlyContinue
	Remove-variable CounterUsersADAccountUpdateFailed -ErrorAction SilentlyContinue
	Remove-variable CounterUsersToSynchronize -ErrorAction SilentlyContinue
	Remove-variable CounterUsersNativeSynchronizationSuccess -ErrorAction SilentlyContinue
	Remove-variable CounterUsersNativeSynchronizationFailed -ErrorAction SilentlyContinue
	Remove-variable CounterUsersNativeSynchronizationNoModification -ErrorAction SilentlyContinue
	Remove-Variable UsersWithNativeSynchonizationError -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedSynchronizationSuccess -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedSynchronizationFailed -ErrorAction SilentlyContinue
	Remove-variable CounterUsersAdvancedSynchronizationNoModification -ErrorAction SilentlyContinue
	Remove-variable CounterUsersDeletionFailed -ErrorAction SilentlyContinue
	Remove-variable CounterUsersDeletionSuccess -ErrorAction SilentlyContinue
	remove-variable	UsersWithDomainUnreachable -ErrorAction SilentlyContinue
	Remove-variable UsersNotFound -ErrorAction SilentlyContinue
	Remove-variable UsersWithADAccountUpdateError -ErrorAction SilentlyContinue
	Remove-Variable UsersWithAdvancedSynchonizationError -ErrorAction SilentlyContinue
	[GC]::Collect()
}
Remove-variable DomainReachable -ErrorAction SilentlyContinue
Remove-variable DomainUnReachable -ErrorAction SilentlyContinue	
$EndTime = (Get-Date)
Write-host "****************************************************************************" -Fore Yellow
Write-Host "|-> the process took " (New-TimeSpan $StartTime $EndTime) -Fore White
If($SendMail -eq $true)
{
	#Creating a global html report then adding reports by site
 	$duration = [math]::round($(($EndTime-$StartTime).totalminutes),2)
	$CompleteJobsummary = "
	
	<p>SharePoint profile synchronization has completed with "

	$JobResult = ""
	$GlobalResult.GetEnumerator() | Sort-Object Name | % {
	
		If ($_.value -le 25)
		{
			$JobResult = "errors (Synchronization success < 25%)"
		}
		ElseIf (($_.value -le 75) -and (($jobResult -eq "") -or ($JobResult -eq "success (Synchronization success > 75%)")))
		{
			$JobResult = "warnings (Synchronization success < 75%)"
		}
		ElseIf ($jobResult -eq "")
		{
			$JobResult = "success (Synchronization success > 75%)"
		}		
	}
	$jobStatus = "SharePoint profile synchronization has completed with " + $jobResult + " on server " + (gc env:computername)	

	$CompleteJobsummary += $jobResult + ".</p>	
	<p><strong><u>Summary:</u></strong></p>

	<table border='1' cellpadding='0' cellspacing='0'>
		<tbody>
			<tr>
				<td>
				<p>Server</p>
				</td>
				<td>
				<p>" + (gc env:computername) + "</p>
				</td>
			</tr>
			<tr>
			<td>
			<p>Sites to synchronize</p>
			</td>
			<td>"
	$GlobalResult.GetEnumerator() | Sort-Object Name | % {
	
		$CompleteJobsummary += "<font color='black'>" + $_.key + " :</font>"
		
		If ($_.value -ge 75)
		{$CompleteJobsummary += "<font color='green'>" + $_.value + " % succeed</font><br />"}
		ElseIf ($_.value -le 25)
		{$CompleteJobsummary += "<font color='red'>" + $_.value + " % succeed</font><br />"}
		Else
		{$CompleteJobsummary += "<font color='orange'>" + $_.value + " % succeed</font><br />"}
	}
	$CompleteJobsummary += "
				</td>				
			</tr>	
			<tr>
				<td>
				<p>Duration</p>
				</td>
				<td>
				<p>$duration minutes (Start:$StartTime / End:$EndTime)</p>
				</td>
			</tr>
		</tbody>
	</table>

	<p>&nbsp;</p>

	<p><strong><u>Details:</u></strong></p>"

	$CompleteJobsummary += $jobsummary

	Send-Email
}
Remove-variable GlobalResult -ErrorAction SilentlyContinue
Remove-variable JobResult -ErrorAction SilentlyContinue
Remove-variable jobsummary -ErrorAction SilentlyContinue
Remove-variable CompleteJobsummary -ErrorAction SilentlyContinue	
if ($TranscriptStatus -ne "Error")
{
Stop-Transcript
}