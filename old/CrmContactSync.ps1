#############################################
# ContactSync.ps1
# 24-Jan-2017
# Yuzuru Kenyoshi
#############################################

get-pssession | remove-pssession

#############################################
## Prepare Exchange			            #####
#############################################

$Office365Username = "admin@aden.partner.onmschina.cn"
$Office365Password = "All.007!"

$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365Password -Force    
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://partner.outlook.cn/PowerShell?proxyMethod=RPS" -Credential $Office365Credentials -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null

#############################################
## Prepare Log File						  ###
#############################################

$batchNo =  Get-Date -Format 'yyyy年MM月dd日HH时mm分'
$historyLog = "c:\temp\ContactSyncLog_"+ $batchNo + ".txt"
$errorLog = "c:\temp\ContactSyncErrorLog_"+ $batchNo + ".txt"

##############################################
function SyncSales(){

	$query = "select mail from humres where blocked=0 and mail is not null" 
	$connectionString = "Data Source=172.16.4.103;Initial Catalog=adencrm;User Id=adensa;Password=123456"

	$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

	$connection.ConnectionString = $connectionString
	$command = $connection.CreateCommand()
	$command.CommandText = $query

	$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command                           
	$dataset = New-Object -TypeName System.Data.DataSet
	$adapter.Fill($dataset)
	$table=$dataset.Tables[0]

	$salesMembers=@()
	$table.Rows.count


	if($table.Rows.count -Gt 1){
		"OK"
		for($i=0;$i -le $table.Rows.count-1;$i++){

			$email = $table.Rows[$i].ItemArray[0].Trim().ToLower()

			if($email.Contains("@adenservices.com")){


			
				$mailBox = get-mailbox $email

				if($mailBox.count -ne 0){
					$salesMembers+=$email
					"Sales;"+ $i + ";" + $email + ";success" >> $historyLog
				}
				else{
					"Sales;"+ $i + ";" + $email + ";nomailbox" >> $errorLog
				}

			}
			else{
				"Sales;"+ $i + ";" + $email + ";noadendomain" >> $errorLog
			}

		}
	
		$salesMembers = $salesMembers | select -uniq
		$error[0] = $null
		Update-DistributionGroupMember -Identity jianji-crm-sales -Member $salesMembers -Confirm:$false
	
		if($error[0] -ne $null){
			"Update Sales Group;" + $error[0] >>  $historyLog
		}
		else{
			"Update Sales Group;Successful" >>  $historyLog
		}
	}

}
function SyncCustomers(){

	$query = "select cnt_email from cicntp where cnt_email is not null" 
	$connectionString = "Data Source=172.16.4.103;Initial Catalog=adencrm;User Id=adensa1;Password=123456"

	$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

	$connection.ConnectionString = $connectionString
	$command = $connection.CreateCommand()
	$command.CommandText = $query

	$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command                           
	$dataset = New-Object -TypeName System.Data.DataSet
	$adapter.Fill($dataset)
	$table=$dataset.Tables[0]

	$customerMembers=@()
	$table.Rows.count
	if($table.Rows.count -Gt 1){
		"OK"
		for($i=0;$i -le $table.Rows.count-1;$i++){
			$i
			$email = $table.Rows[$i].ItemArray[0].Trim().ToLower()

			if($email.Contains("@adenservices.com")){
				######  Do Nothing
				"Customer;"+ $i + ";" + $email + ";AdenEmail" >>　$errorLog
			}
	
			else{
				######  非AdenService邮件
				$mailContact = Get-MailContact -Identity $email
				$mailContactExists= $?

				if($mailContactExists -eq $True){
				
					$customerMembers+= $email
					"Customer;"+ $i + ";" + $email + ";OK" >> $historyLog
				}
				else{
					$error[0] = $null
					New-MailContact -Name $email -ExternalEmailAddress $email
					Set-MailContact -Identity $email -HiddenFromAddressListsEnabled $true

					if($error[0] -eq $null){
						$customerMembers+= $email
						"Customer;"+ $i + ";" + $email + ";CreateMailContactSucess" >> $historyLog
					}
					else{
						"Customer;"+ $i + ";" + $email + ";" +$error[0] >>　$errorLog
					}
				}
			}

		
		}

		$customerMembers = $customerMembers | select -uniq
		$error[0]=$null
		Update-DistributionGroupMember -Identity jianji-crm-customers -Member $customerMembers -Confirm:$false
		if($error[0] -ne $null){
			"Update Customer Group;" + $error[0] >>  $historyLog
		}
		else{
			"Update Customer Group;Successful" >>  $historyLog
		}

	}


}

SyncSales
SyncCustomers

Remove-PSSession $exchangeSession
