#############################################
# ReportToSync.ps1
# 1-Aug-2017
# Author: Winter Wang
# ver: 1.0
#
# This tool is to update field "Manager" in Active Directory. The informaiton is coming from BenQ outlook view.
#############################################

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

$query = "select * from [dbo].[v_OutlookData] where LeaveDate = ''" 
$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"

#############################################
## Prepare Log File						#####
#############################################

$batchNo =  Get-Date -Format 'yyyy-MM-dd_hh'
$syncLog = "c:\temp\BenQADSync\ReportToSync_"+ $batchNo + ".txt"

#############################################

function Main {
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
		# read employeeID and ReportToID
		if($myvalue -eq "employeeID"){$employeeIDIndex = $i-1		}
		if($myvalue -eq "ReportToID"){$ReportToIDIndex = $i-1		}
	}

	for($i=1;$i -le $table.Rows.count;$i++){
		$i
		$uData = @{}
		$uData.employeeNum = HandleNull($table.Rows[$i-1].ItemArray[$employeeIDIndex])
		$uData.ReportToID = HandleNull($table.Rows[$i-1].ItemArray[$ReportToIDIndex])	
		
		# Get employeeID and ReportToMail from DB.
		$employeeID = $uData.employeeNum
		$ManagerID = $uData.ReportToID
		
		# Check if ReportToMail is empty
		if ($ManagerID -ne $null) {
		
			# Check if Manager is same one in AD and BenQ
			$ADmgrAccount = Get-ADUser -Filter {employeeID -eq $employeeID} -Properties manager | select manager

			if ($ADmgrAccount.manager -eq $null) {
				$SetMgrAction = "yes"
				}
			else {
				$ADmgrID = get-aduser -identity $ADmgrAccount.manager -properties employeeID | select employeeID
				if ($ADmgrID.employeeID -ne $ManagerID) {
					$SetMgrAction = "yes"
					}
				else {
					$SetMgrAction = "no"
					}
				}
				
			if ($SetMgrAction -eq "yes") {
				# Get mgr samAccountName
				$adAccount = Get-ADUser -Filter {employeeID -eq $employeeID} | Select-Object samAccountName
				$mgrAccount = Get-ADUser -Filter {employeeID -eq $ManagerID} | Select-Object samAccountName
				$sam = $adAccount.SamAccountName
				$mgr = $mgrAccount.SamAccountName
				
				if($sam -ne $null){
				#	if ($employeeID -eq "600000070") {
					Set-ADUser $sam -Manager $mgr
					# Log
					$uData.employeeNum + ";" + $sam + ";" + $mgr  >> $Synclog	
				#	}
				}	
			}
		}	
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

main
