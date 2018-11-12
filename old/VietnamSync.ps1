#
# VietnamSync.ps1
#

get-pssession | remove-pssession


#############################################
## Prepare Exchange and Office365       #####
#############################################


$Office365Username = "admin@aden.partner.onmschina.cn"
$Office365Password = "All.007!"

$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365Password -Force    
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://partner.outlook.cn/PowerShell?proxyMethod=RPS" -Credential $Office365Credentials -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null
Connect-MsolService -Credential $Office365Credentials


#############################################
## Prepare Sharepoint Online Connection #####
#############################################

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
$siteUrl = "https://aden.sharepoint.cn/sites/globalit/vietnam/"
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$accountName = "office365@adenservices.com"  ## This should not be partner.onmschina ID
$password = ConvertTo-SecureString -AsPlainText -Force "Pass@189"
$spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($accountName, $password)
$ctx.Credentials = $spCredentials

$vietnamList = $ctx.Web.Lists.GetByTitle("VietNam Employees List")

#############################################
## Prepare Log File						#####
#############################################

$batchNo =  Get-Date -Format 'yyyy年MM月dd日HH时mm分'
$historyLog = "c:\temp\VietnamSync\VietnamSyncLog_"+ $batchNo + ".txt"
$vietnamNewAccountLog = "c:\temp\VietnamSync\VietnamNewAccountLog.txt"

##############################################
##############################################


function Main(){
	
	$users=GetListItems

	for($i=0;$i -lt $users.length; $i++){

		$udata = $users[$i]
		$email = $udata["EMAIL"]

		if($udata["OUTDATE"] -ne $null){
			$hide = $True 
		}
		else{
			$hide = $False
		}

		Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue -ErrorVariable errorVariable
		$errorVariable
		$email
		if($errorVariable -ne $null){
			if($hide -eq $false){
				"new"
				New_O365Account($udata);
				Update_O365Account($udata);			
			}  
		}
		else{
			if($hide -eq $false){
				"update"
				Update_O365Account($udata);
			}
			else{
				"hide"
				Hide_O365Account($udata);
			}			
		}
	}
}

function New_O365Account(){
	$email = $udata["EMAIL"]
	$displayname = $udata['DISPLAYNAME'].replace("`r`n","`n")

	$region = $udata['REGION']
	if($region -eq'Viet Nam'){$region = 'VN'}
	elseif($region -eq 'Indonesia'){$region = 'ID'}
	elseif($region -eq 'LAO'){$region = 'LA'}

    try{
		New-MsolUser -UserPrincipalName $email `
			-DisplayName $displayname `
			-UsageLocation $region `
			-Password "Aden@123"

	

        $email +";CreateNewO365;Success;"+ $batchNo >> $vietnamNewAccountLog
		
		###Mail###
		$subject = "New Vietnam Office365 Account"
		$body = "<p>Email:" + $email + "</p>"
		$body = $body + "<p>Region:" + $region + "</p>"
		$body = $body + "<p>Password:Aden@123</p>"
		$recepients = @("888it@adenservies.com","yuzuru.kenyoshi@189csp.com","yuanjie@189csp.com","hoang.hai@adenservices.com")
		#$recepients = @("yuzuru.kenyoshi@189csp.com")
		Send-MailMessage -To $recepients -Subject $subject -Body $body -UseSsl -Port 587  -SmtpServer 'smtp.partner.outlook.cn' -From "admin@aden.partner.onmschina.cn" -BodyAsHtml -Credential $Office365Credentials -Encoding UTF8
	    Set-MsolUserLicense -userPrincipalName $email -AddLicenses "reseller-account:EXCHANGESTANDARD"
             
    }
    catch{
        $email +";CreateNewO365;failed-"+$error[0] +";" + $batchNo >> $vietnamNewAccountLog
        $error[0] = $null
    }

}

function Update_O365Account(){

	$email = $udata['EMAIL']
	$displayname = $udata['DISPLAYNAME'].replace("`r`n","`n")
	$firstname =$udata['FIRSTNAME']
	$lastname =$udata['LASTNAME']
	$departname =$udata['DEPT']
	$office =$udata['OFFICE']
	$title =$udata['JOBTITLE']
	$mbphone=$udata['MOBILEPHONE']
	$offphone =$udata['OFFICEPHONE']
	$initials = $udata['ABBREVIATION']
	$address = $udata['COMPANYADDRESS']
	$region = $udata['REGION']

	if($region -eq'Viet Nam'){
		$region = 'VN'
	}
	elseif($region -eq 'Indonesia'){
		$region = 'ID'
	}
	elseif($region -eq 'LAO'){
		$region = 'LA'
	}
             
	Set-MsolUser -UserPrincipalName $email `
				-DisplayName $displayname `
				-FirstName $firstname `
				-LastName $lastname `
				-Department $departname `
				-Office $office `
				-Title $title `
				-StreetAddress $address `
				-UsageLocation $region `
				-MobilePhone $mbphone `
				-PhoneNumber $offphone
			
	Set-User -Identity $email -Initials $initials
	Set-Mailbox -Identity $email -HiddenFromAddressListsEnabled $false
		
	### Set Lisence	
	$user = Get-MsolUser -UserPrincipalName $email 
	if($user.IsLicensed -ne $True){
		Set-MsolUserLicense -userPrincipalName $user.userPrincipalName -AddLicenses "reseller-account:EXCHANGESTANDARD"  
	}

	### Add Alias
	$alias = $email.replace("@adenservices.com","@aden.partner.onmschina.cn") 
	Set-Mailbox -Identity $email -EmailAddresses @{add=$alias}
}

function Hide_O365Account($udata){
	$email = $udata["EMAIL"]
	Set-Mailbox -Identity $email -HiddenFromAddressListsEnabled $true
}


function GetListItems(){
	$camlQuery = new-object Microsoft.SharePoint.Client.CamlQuery;
	$camlQuery.ViewXml = "<View>
		<Query>
			<OrderBy> 		  
				<FieldRef Name='ID' Ascending='FALSE'/> 
			</OrderBy>
			<RowLimit>limit</RowLimit>
		</Query>
	</View>"
	$items = $vietnamList.GetItems($camlQuery)
	$ctx.Load($items)
	$ctx.ExecuteQuery();
	if($items -ne $null){
		return $items
	}
	else{
		return $null
	}
}
function UploadLogFile(){

	$Library = $ctx.Web.Lists.GetByTitle("Documents")
	$ctx.Load($Library)
	$ctx.ExecuteQuery()

	##################
	$FileStream = New-Object IO.FileStream($historyLog,[System.IO.FileMode]::Open)
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = "VietnamSync_Latest.txt"

	$Upload = $Library.RootFolder.Files.Add($FileCreationInfo)
	$ctx.Load($Upload)
	$ctx.ExecuteQuery()
	##################

	$FileStream = New-Object IO.FileStream($vietnamNewAccountLog,[System.IO.FileMode]::Open)
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = "VietnamNewAccountLog.txt"

	$Upload = $Library.RootFolder.Files.Add($FileCreationInfo)
	$ctx.Load($Upload)
	$ctx.ExecuteQuery()

}

Main
UploadLogFile
Remove-PSSession $exchangeSession



