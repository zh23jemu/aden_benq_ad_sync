#
# Retir_Forward.ps1
#

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

$query = "select * from [dbo].[v_OutlookData] where forwardMail<>'' and OutDate< GETDATE() and forwardDate> GETDATE()"
$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"

#DATE_ADD(NOW(), INTERVAL -1 MONTH)

#############################################
## Prepare Log                         #####
#############################################
$batchNo =  Get-Date -Format 'yyyyMMdd'
$logPath = "C:\temp\Forward_Convert\"+ $batchNo +".txt"
"Email;OutDate;ForwardMail;ForwardDate" >> $logPath

# Prepare office 365
$Office365Username = "admin@aden.partner.onmschina.cn"
$Office365Password = "All.007!"

$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365Password -Force    
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://partner.outlook.cn/PowerShell?proxyMethod=RPS" -Credential $Office365Credentials -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null

Connect-MsolService -Credential $Office365Credentials -AzureEnvironment AzureChinaCloud
################

function Main {

	Write-Verbose 'in SQL Server mode'
	$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

	$connection.ConnectionString = $connectionString
	$command = $connection.CreateCommand()
	$command.CommandTimeout=0
	$command.CommandText = $query

	$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
	$dataset = New-Object -TypeName System.Data.DataSet
	$adapter.Fill($dataset)
	$table=$dataset.Tables[0]
	
	for($i=0;$i -lt $table.Columns.count;$i++){
		$myvalue = $table.Columns[$i].toString()

		if($myvalue -eq "Email"){		$EmailIndex = $i		}
		if($myvalue -eq "OutDate"){		$OutDateIndex = $i		}
		if($myvalue -eq "ForwardMail"){		$ForwardMailIndex = $i		}
		if($myvalue -eq "ForwardDate"){ $ForwardDateIndex = $i	 }
	}

	for($i=0;$i -lt $table.Rows.count;$i++){

		$i
		#$table.Rows
		$uData = @{}

		$uData.email = HandleNull($table.Rows[$i].ItemArray[$EmailIndex])
		$uData.outDate = HandleNull($table.Rows[$i].ItemArray[$OutDateIndex])		
		$udata.forwardMail = HandleNull($table.Rows[$i].ItemArray[$ForwardMailIndex])
		$udata.forwardDate = HandleNull($table.Rows[$i].ItemArray[$ForwardDateIndex])

		#if($udata.ForwardDate-eq $null){
		#	"����Ϊ��"
		# return;
		#}

		
		#$uData.email+";"+$uData.outDate+";"+$uData.forwardMail+";"+$uData.forwardDate >> $logPath
	    Forward($uData)
    }
}
function Forward($uData){
		$email = $uData.email
	$msolUser = Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue
	If ($msolUser -ne $Null) { 
		#$msolUser.IsLicensed
		$uData.email+";"+$uData.outDate+";"+$uData.forwardMail+";"+$uData.forwardDate+";"+$msolUser.IsLicensed
		$uData.email+";"+$uData.outDate+";"+$uData.forwardMail+";"+$uData.forwardDate+";"+$msolUser.IsLicensed >> $logPath
    Get-Mailbox $email | Where {$_.ForwardingSMTPAddress -ne $null} | Set-Mailbox -ForwardingSMTPAddress $uData.forwardMail -DeliverToMailboxAndForward $false
	#Set-Mailbox $email -ForwardingSMTPAddress  "admin@aden.partner.onmschina.cn" -DeliverToMailboxAndForward $False 
	#Get-Mailbox $email| Where {$_.RecipientTypeDetails -ne "SharedMailbox"} | Set-Mailbox $email -Type shared
	#Set-MailboxAutoReplyConfiguration -Identity $email -AutoReplyState Enabled -ExternalMessage "Ա������ְ" -InternalMessage "Ա������ְ"
	$user = Get-MsolUser -UserPrincipalName $email 

	$forwardmail = Get-Mailbox $email
    $forward = $sharedmail.RecipientTypeDetails

	}
	else{
		$uData.email+";"+$uData.outDate+";"+$uData.forwardMail+";"+$uData.forwardDate+";NoO365"
		$uData.email+";"+$uData.outDate+";"+$uData.forwardMail+";"+$uData.forwardDate+";NoO365" >> $logPath	
		
	}

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