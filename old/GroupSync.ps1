#
# GroupSync.ps1
#
#
# Retir_Forward.ps1
#

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################
$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"
$query = "select DISTINCT emailgroup from v_UserDL" 


#SELECT DISTINCT 列名称 FROM 表名称

$queryEmail = "select * from v_UserDL" 

#############################################




#######################################
#365
#######################################
	$Office365Username = "admin@aden.partner.onmschina.cn"
	$Office365Password = "All.007!"
	#$Office365Password = "Aden@123"
$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365Password -Force    
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://partner.outlook.cn/PowerShell?proxyMethod=RPS" -Credential $Office365Credentials -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null

######################################################
#############################################
## Prepare Log File						  ###
#############################################

$batchNo =  Get-Date -Format 'yyyy年MM月dd日HH时mm分'
############################              
	$members= @{}
function Main(){
	CheckGroups
	UpdateGroupsMembers
}
function CheckGroups {

	Write-Verbose 'in SQL Server mode'
	$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

	$connection.ConnectionString = $connectionString
	$command = $connection.CreateCommand()
	$command.CommandText = $query
	$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
	$dataset = New-Object -TypeName System.Data.DataSet
	$adapter.Fill($dataset)
	$table=$dataset.Tables[0]
	$emailgroups = New-Object -TypeName System.Collections.ArrayList
	for($i=0;$i -lt $table.Rows.count;$i++){

		$emailgroupname = HandleNull($table.Rows[$i].ItemArray[0])
		$emailgroups.add($emailgroupname)
    }


	$o365Groups =  Get-DistributionGroup
	for($i=0;$i -lt $emailgroups.count;$i++){
		if($o365Groups -contains $emailgroups[$i]){
		"组存在" +$emailgroups[$i]
		}
	    else {
	   "组不存在" + $emailgroups[$i]
		CreateGroup($emailgroups[$i])

		}
		$members[$emailgroups[$i]] = New-Object -TypeName System.Collections.ArrayList


	}



}
function UpdateGroupsMembers(){
	
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

	$connection.ConnectionString = $connectionString
	$command = $connection.CreateCommand()
	$command.CommandText = $queryEmail
	$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
	$dataset = New-Object -TypeName System.Data.DataSet
	$adapter.Fill($dataset)
	$table=$dataset.Tables[0]
	for($i=0;$i -lt $table.Rows.count;$i++){
	
	    $email = HandleNull($table.Rows[$i].ItemArray[0])
		$emailgroupname = HandleNull($table.Rows[$i].ItemArray[1])


#		$email
#		$emailgroupname


			$mailBox = get-mailbox $email

			if($mailBox.count -ne 0){
				$members[$emailgroupname].add($email)
						$logPath = "C:\temp\groupSync\" +$emailgroupname+"_"+ $batchNo+".txt"
		                $email>>$logPath
			}
			else{
						$logPath = "C:\temp\groupSync\" +$emailgroupname+"_Error_"+ $batchNo+".txt"
		               $email>>$logPath
			}




	
		

		
	}


	foreach ($key in $members.keys){
		$groupName = $key 
		$emails = $members[$key]
		$emails = $emails | select -uniq
		Update-DistributionGroupMember -Identity $groupName -Member $emails -Confirm:$false

		
	}

}





function BenqGroupSync($groupname){
		Write-Verbose 'in SQL Server mode'
	$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

	$connection.ConnectionString = $connectionString
	$command = $connection.CreateCommand()
	$command.CommandText = $queryEmail
	$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
	$dataset = New-Object -TypeName System.Data.DataSet
	$adapter.Fill($dataset)
	$table=$dataset.Tables[0]
	$emailArray = New-Object -TypeName System.Collections.ArrayList

	$info = @{}
	for($i=1;$i -le $table.Rows.count;$i++){
	
	    $email = HandleNull($table.Rows[$i-1].ItemArray[0])
		$emailgroupname = HandleNull($table.Rows[$i-1].ItemArray[1])
		if($groupname -contains $emailgroupname){
			$emailArray.add($email)
			$info.$emailgroupname=$emailArray
			$info
		}
	}
}

function CreateGroup($groupname){
	"创建"+ $groupname
	#New-DistributionGroup $groupname
	$primarySmtpAddress = $groupname + "@adenservices.com"
	New-DistributionGroup $groupname -PrimarySmtpAddress $primarySmtpAddress

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