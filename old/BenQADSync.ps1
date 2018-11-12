#############################################
# BenQADSync.ps1
# 28-Feb-2017
# Yuzuru Kenyoshi
# This does not include License Assignment
#############################################

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

$query = "select * from [dbo].[v_OutlookData]" 
$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"

#############################################
## Prepare Log File						#####
#############################################

$batchNo =  Get-Date -Format 'yyyy年MM月dd日HH时mm分'
$historyLog = "c:\temp\BenQADSync\BenqSyncLog_"+ $batchNo + ".txt"

"Update: "+$batchNo >> $historyLog
"Email;" + "Employee Id;"  + "Display Name;"  + "First Name;"  + "Last Name;"+ "Initials;"+ "Job Title;"+ "Dept;"  + "Region;" + "Full Office Phone;" + "Mobile;" + "InDate;"+ "OutDate;"  + "Hide;" + "Type;" + "Log" >> $historyLog

$changedEmailLog = "c:\temp\BenqADSync\ChangedEmailLog_Sync\ChangedEmailLog_"+ $batchNo + ".txt"


"Update: "+$batchNo >> $changedEmailLog
"EmmployeeEd;" + "BenQEmail;"   + "SamAccountName;"+ "AdUserPrincipalName;" >>$changedEmailLog


$createAdLog = "c:\temp\BenQADSync\CreateADLog.txt"
$createAdErrorLog = "c:\temp\BenQADSync\CreateADErrorLog_"+ $batchNo + ".txt"
#######################################
#365
#######################################
	$Office365Username = "admin@aden.partner.onmschina.cn"
	$Office365Password = "All.007!"
$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365Password -Force    
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

#$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://partner.outlook.cn/PowerShell?proxyMethod=RPS" -Credential $Office365Credentials -Authentication "Basic" -AllowRedirection
#Import-PSSession $exchangeSession -AllowClobber | Out-Null


#############################################
## Sharepoint    						#####
#############################################

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
$siteUrl = "https://aden.sharepoint.cn/sites/globalit/usermanagement/"
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$accountName = "mark.o365@adenservices.com"
$password = ConvertTo-SecureString -AsPlainText -Force "Hello189"
$spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($accountName, $password)
$ctx.Credentials = $spCredentials

$ouList = $ctx.Web.Lists.GetByTitle("benq-ou")

#############################################

function Main {
    'entering main'
	Write-Verbose 'in SQL Server mode'
	$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

	$connection.ConnectionString = $connectionString
	$command = $connection.CreateCommand()
	$command.CommandText = $query
	$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
	$dataset = New-Object -TypeName System.Data.DataSet
	$adapter.Fill($dataset)
	$table=$dataset.Tables[0]
	
	for($i=1;$i -le $table.Columns.count;$i++){
		$myvalue = $table.Columns[$i-1].toString()
		$myvalue

		if($myvalue -eq "EMPLOYEEID"){		 $EMPLOYEEIDIndex = $i-1		}
		if($myvalue -eq "FIRSTNAME" ){		 $FIRSTNAMEIndex = $i-1		}
		if($myvalue -eq "LASTNAME"){		 $LASTNAMEIndex = $i-1		}
		if($myvalue -eq "DisplayName"){		 $DisplayNameIndex = $i-1		}
		if($myvalue -eq "Initials"){		 $InitialsIndex = $i-1		}
		if($myvalue -eq "Email"){		 　　$EmailIndex = $i-1		}
		if($myvalue -eq "Dept"){		 $DeptIndex = $i-1		}
		if($myvalue -eq "Region"){		 $RegionIndex = $i-1		}
		if($myvalue -eq "JobTitle"){		 $JobTitleIndex = $i-1		}
		if($myvalue -eq "OFFICEPHONE"){		 $OFFICEPHONEIndex = $i-1		}
		if($myvalue -eq "EXT"){		 $EXTIndex = $i-1		}
		if($myvalue -eq "MobilePhone"){		 $MobilePhoneIndex = $i-1		}
		if($myvalue -eq "InDate"){		 $InDateIndex = $i-1		}
		if($myvalue -eq "OutDate"){		 $OutDateIndex = $i-1		}
	}

	for($i=1;$i -le $table.Rows.count;$i++){

		#marked by billy
        #$i

		$uData = @{}
		$uData.employeeNum = HandleNull($table.Rows[$i-1].ItemArray[$EMPLOYEEIDIndex])
		$uData.firstName = HandleNull($table.Rows[$i-1].ItemArray[$FIRSTNAMEIndex])
		$uData.lastName = HandleNull($table.Rows[$i-1].ItemArray[$LASTNAMEIndex])
		$uData.displayName = HandleNull($table.Rows[$i-1].ItemArray[$DisplayNameIndex])
		$uData.initials = HandleNull($table.Rows[$i-1].ItemArray[$InitialsIndex])	
		$uData.email = HandleNull($table.Rows[$i-1].ItemArray[$EmailIndex])
		$uData.deptName = HandleNull($table.Rows[$i-1].ItemArray[$DeptIndex])
		$uData.regionName = HandleNull($table.Rows[$i-1].ItemArray[$RegionIndex])
		$uData.jobTitle = HandleNull($table.Rows[$i-1].ItemArray[$JobTitleIndex])
		$uData.officeNum = HandleNull($table.Rows[$i-1].ItemArray[$OFFICEPHONEIndex])
		$uData.ext = HandleNull($table.Rows[$i-1].ItemArray[$EXTIndex])
		$uData.mobileNum = HandleNull($table.Rows[$i-1].ItemArray[$MobilePhoneIndex])
		$uData.inDate = HandleNull($table.Rows[$i-1].ItemArray[$InDateIndex])
		$uData.outDate = HandleNull($table.Rows[$i-1].ItemArray[$OutDateIndex])
		

		if(($uData.jobTitle -eq $null -and $uData.regionName -eq $null) -or $uData.outDate -ne $null){
			$uData.hide = $True 
		}
		else{
			$uData.hide = $False
		}

		$uData.fullOfficePhone = FullOfficePhone($uData)
		$employeeId = $uData.employeeNum
		$adAccount = Get-ADUser -Filter {EmployeeID -eq $employeeId} | Select-Object samAccountName,UserPrincipalName
		#$adAccount
		$sam = $adAccount.SamAccountName
		#$sam
		if($sam -ne $null){
			UpdateUser($uData)
		}
		else{
			if($uData.hide -eq $True){
				$uData.type = 'Do Nothing'							
			}
			else{
				#"Email:" + $uData.email
				CreateAD($uData)	
			}
		}


		CreateLog($uData)			
	}
	UploadLogFile
}

function UpdateUser($uData){


	# User Found in AD
	if($uData.hide -eq $True){
		$uData.type = 'AD Hide'
		Set-ADUser $sam -Replace @{msExchHideFromAddressLists=$True} 
	}
	else{
		#$adAccount.UserPrincipalName
		if($uData.email -ne $adAccount.UserPrincipalName){
			$uData.employeeNum+";"+ $uData.email +";"+ $adAccount.SamAccountName +";"+ $adAccount.UserPrincipalName >>$changedEmailLog
			$uData.type = 'BenQ Email Mismatch'
		}	
		else{
            #$sam
			$uData.type = 'AD Sync'

			Set-ADUser $sam `
			-SurName $uData.lastName `
			-GivenName $uData.firstName `
			-DisplayName $uData.displayName `
			-Department $uData.deptName `
			-Office $uData.regionName `
			-Title $uData.jobTitle `
			-OfficePhone $uData.fullOfficePhone `
			-MobilePhone $uData.mobileNum `
			-Initials $uData.initials `
			-Replace @{msExchHideFromAddressLists=$False}

		}				
	} 

}

function CreateAD($uData){

	$uData.type = 'AD Not Found'
		
	$log = @{}

	$log.email = ""
	$log.regionDept = ""
	$log.path = ""
	$log.result = ""	

	##### Check #####
    $email = $uData.email
    $name =  $email.substring(0,$email.IndexOf("@")).Trim() # Modified by billy 20180627
	#$name  = $uData.email.replace("@adenservices.com","").Trim()
	$useraccount = Get-ADUser -Filter {sAMAccountName  -eq $name}
	$useraccount2 = Get-ADUser -Filter {userPrincipalName  -eq $email }
	if($useraccount -ne $Null){
		$log.result = "SAMACCOUNT已经存在"
		$udata.email +";" + $log.result>> $createAdErrorLog
	}
	elseif($useraccount2 -ne $Null){
		$log.result = "UPN已经存在"	
		$udata.email +";" + $log.result>> $createAdErrorLog		
	}
	else{
	#不存在 新建
		$employeeNum = HandleNull($table.Rows[$i-1].ItemArray[$EMPLOYEEIDIndex])
		$firstName = HandleNull($table.Rows[$i-1].ItemArray[$FIRSTNAMEIndex])
		$lastName = HandleNull($table.Rows[$i-1].ItemArray[$LASTNAMEIndex])
		$disName = HandleNull($table.Rows[$i-1].ItemArray[$DisplayNameIndex])
		$initialsName = $table.Rows[$i-1].ItemArray[$InitialsIndex]
		$deptName = HandleNull($table.Rows[$i-1].ItemArray[$DeptIndex])
		$regionName = HandleNull($table.Rows[$i-1].ItemArray[$RegionIndex])
		$jobTitle = HandleNull($table.Rows[$i-1].ItemArray[$JobTitleIndex])
		$officeNum = HandleNull($table.Rows[$i-1].ItemArray[$OFFICEPHONEIndex])
		$ext = HandleNull($table.Rows[$i-1].ItemArray[$EXTIndex])
		$mobileNum = HandleNull($table.Rows[$i-1].ItemArray[$MobilePhoneIndex])
		
		################


		$log.regionDept = $regionName + "-" + $deptName

		$camlQuery = new-object Microsoft.SharePoint.Client.CamlQuery;
		$camlQuery.ViewXml = "<View>
		<Query>
			<Where>
				<And>
					<Eq>   
					<FieldRef Name='Region' /> 
					<Value Type='Text'>"+$regionName+"</Value>
					</Eq>
					<Eq>   
					<FieldRef Name='Dept' /> 
					<Value Type='Text'>"+$deptName+"</Value>
					</Eq>
				</And>
			</Where>
		</Query>
		</View>"
        $items = @() #added by billy
		$items = $ouList.GetItems($camlQuery)
		$ctx.Load($items)
		$ctx.ExecuteQuery();

<#		if($items -ne $null){
			$item = $items[0]
			$ouPath = ""
			if($item["String5"] -ne $null){$ouPath = "OU="+ $item["String5"] + "," }
			if($item["String4"] -ne $null){$ouPath = $ouPath + "OU="+ $item["String4"]+ "," }
			if($item["String3"] -ne $null){$ouPath = $ouPath + "OU="+ $item["String3"]+ "," }
			if($item["String2"] -ne $null){$ouPath = $ouPath + "OU="+ $item["String2"]+ "," }
			$ouPath = $ouPath + "OU="+$item["String1"] + ",DC=CHOADEN,DC=COM"
            'oupath is' + $ouPath
			$log.path = $ouPath
		}#>
#		else{
			#OU Info not found in Sharepoint
			$ouPath = "OU=BENQ,OU=ADEN-Users,DC=CHOADEN,DC=COM"
			$log.path = $ouPath
#		}
        
        if([adsi]::Exists("LDAP://$ouPath"))
        {			
            #marked by billy
            "create new user " + $name + ' in ' + $ouPath
            "Log Start">>  "c:\temp\BenQADSync\CreateADLogTemp.txt"		
			New-ADUser $name `
			-SamAccountName $name `
			-userprincipalname $email `
			-EmailAddress $email `
			-EmployeeID $employeeNum `
			-Path $ouPath   `
			-AccountPassword (ConvertTo-SecureString "Aden@123" -AsPlainText -Force) `
			-ChangePasswordAtLogon $false `
			-enabled $true
			$log.result = "创建"
			$udata.email +";" + $log.result + ";" + $batchNo  >> $createAdLog


			$subject = "新BenQ账号提醒"
			$body = "<p>Email:" + $udata.email + "</p>"
			$body = $body + "<p>Region&Dept:" + $log.regionDept + "</p>"
			$body = $body + "<p>INDATE:" + $udata.inDate + "</p>"
			$body = $body + "<p>OU:" + $log.path + "</p>"
			$body = $body + "<p>Result:" + $log.result + "</p>"
			$recepients = @("yuzuru.kenyoshi@189csp.com","yuanjie@189csp.com","888it@adenservices.com","billy.zhou@adenservices.com")
            #Send-MailMessage -To billy.zhou@adenservices.com -Subject test -Body test -UseSsl -Port 25  -SmtpServer 'adenservices-com.mail.protection.partner.outlook.cn' -From "admin@aden.partner.onmschina.cn" -BodyAsHtml -Encoding UTF8			
            #Send-MailMessage -To $recepients -Subject $subject -Body $body -UseSsl -Port 587  -SmtpServer 'smtp.partner.outlook.cn' -From "admin@aden.partner.onmschina.cn" -BodyAsHtml -Credential $Office365Credentials -Encoding UTF8
			


		}
		else{
			$log.result = "指定的OU"+$log.path +"不存在"
			$udata.email +";" + $log.result >> $createAdErrorLog
		}
		
			
	}
	$uData.log = $log.result

	#$udata.email +";" + $uData.log + ";" + $batchNo  >> $createAdLog

}

function CreateLog($uData){
	$uData.email + ";" 	+ $uData.employeeNum +";" 	+ $uData.displayName +";" 	+ $uData.firstName +";" 	+ $uData.lastName +";" 	+ $uData.initials +";" 	+ $uData.jobTitle +";" 	+ $uData.deptName +";" 	+ $uData.regionName + ";"	+ $uData.fullOfficePhone + ";" 	+ $uData.mobileNum + ";" 	+ $uData.inDate +";"	+ $uData.outDate+";"	+ $uData.hide + ";"	+ $uData.type + ";"	+ $uData.log>> $historyLog
}

function UploadLogFile(){


	
	
	##################################


	$Library = $ctx.Web.Lists.GetByTitle("Documents")
	$ctx.Load($Library)
	$ctx.ExecuteQuery()

	###################################

	$FileStream = New-Object IO.FileStream($historyLog,[System.IO.FileMode]::Open)
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = "BenqSync_Latest.txt"

	$Upload = $Library.RootFolder.Files.Add($FileCreationInfo)
	$ctx.Load($Upload)
	$ctx.ExecuteQuery()

	#################################

	$FileStream = New-Object IO.FileStream($createAdLog,[System.IO.FileMode]::Open)
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = "BenqSync_NewAccountLog.txt"

	$Upload = $Library.RootFolder.Files.Add($FileCreationInfo)
	$ctx.Load($Upload)
	$ctx.ExecuteQuery()
	
	#################################

	$FileStream = New-Object IO.FileStream($createAdErrorLog,[System.IO.FileMode]::Open)
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = "BenqSync_CreateADErrorLog_Latest.txt"

	$Upload = $Library.RootFolder.Files.Add($FileCreationInfo)
	$ctx.Load($Upload)
	$ctx.ExecuteQuery()
	#################################

	$FileStream = New-Object IO.FileStream($changedEmailLog,[System.IO.FileMode]::Open)
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = "BenqSync_ChangedEmailLog_Latest.txt"

	$Upload = $Library.RootFolder.Files.Add($FileCreationInfo)
	$ctx.Load($Upload)
	$ctx.ExecuteQuery()

}


function FullOfficePhone($uData){
		if($uData.officeNum -ne $null -and $uData.ext -ne $null){
			if($uData.officeNum -eq "" -or $uData.ext -eq ""){
				$fullOfficePhone = ""
			}
			else{
				$fullOfficePhone = $uData.officeNum+","+$uData.ext				
			}
		}
		elseif($uData.officeNum -eq $null -and $uData.ext -ne $null){
			$fullOfficePhone = $uData.officeNum
		}
		elseif($uData.officeNum -ne $null -and $uData.ext -eq $null){
			$fullOfficePhone = $uData.officeNum
		}
		elseif($uData.officeNum -eq $null -and $uData.ext -eq $null){
			$fullOfficePhone = $null
		}

		return $fullOfficePhone
}



function HandleNull($oldvalue){
	if($oldvalue.Equals([DBNull]::Value)){
		$value = $null
	}
	elseif($oldvalue -eq ''){
		$value =  $null
	}
	else{
		$value = $oldvalue
	}
	return $value

}

Main




#Remove-PSSession $exchangeSession
