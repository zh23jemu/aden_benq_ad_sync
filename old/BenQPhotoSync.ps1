#关闭所有的Session（万一调试中关闭）
get-pssession | remove-pssession

function Invoke-DatabaseQuery {
    [CmdletBinding()]
    param (
        [string]$connectionString,
        [string]$query,
        [switch]$isSQLServer
    )
    if ($isSQLServer) {
        Write-Verbose 'in SQL Server mode'
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    } else {
        Write-Verbose 'in OleDB mode'
        $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    }

    $connection.ConnectionString = $connectionString
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $connection.Open()

	$result= $command.ExecuteReader()
	$table = new-object “System.Data.DataTable”
    $table.Load($result)

	#fileName
	$datetime =  Get-Date -Format 'yyyy-MM-dd hh-mm-ss'
	$extension = ".txt"
	$logFile = "c:\temp\SetPhoto\SetUserPhoto_" + $datetime + $extention
	"Upload Photo Data" >> $logFile

	#fileName_NoPhoto
	$logFileNoPhoto = "c:\temp\SetPhoto\NoPhoto_" + $datetime + $extention
	"Below Users do not have PhotoData" >> $logFileNoPhoto


	for($i=1;$i -le $table.Rows.count;$i++){

			#$ar.add($table.Rows[$i-1].ItemArray[4])
			$email = $table.Rows[$i-1].ItemArray[6]
			#$pt.add($table.Rows[$i-1].ItemArray[13])
			$photoValue = $table.Rows[$i-1].ItemArray[17]
			
			#Prod Env
			$nameEmail  = $email.Trim()
			
			#Test Env
			#$name  = $email.Trim().replace("@adenservices.com","")
			#$nameEmail = $name.Trim() + "@adenservice.to365.org"

            if($photoValue.Equals([DBNull]::Value)){
				$email >> $logFileNoPhoto
			}
			else{
				AddPhoto $nameEmail $photoValue $logFile
			}   

	}
    $connection.close()
}


function CheckO365($email){
	Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue -ErrorVariable errorVariable
	if($errorVariable -ne $null){  
		return $false
	}
	else{
		return $true
	}
}

function AddPhoto($email,$pv,$logFile){
	$result = CheckO365($email)
	if($result -eq $true){
		$email >> $logFile
		set-userPhoto -Identity $email -PictureData ($pv) -Confirm:$false
	}
}


Get-PSSession

#Test Env
#$Office365Username = "mark@adenservice.to365.org"
#$Office365Password = "pass@word1"

#Prod Env
$Office365Username = "admin@aden.partner.onmschina.cn"
$Office365Password = "All.007!"

$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365Password -Force    
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://partner.outlook.cn/PowerShell?proxyMethod=RPS" -Credential $Office365Credentials -Authentication "Basic" -AllowRedirection
Import-PSSession  $exchangeSession -DisableNameChecking
Connect-MsolService -Credential $Office365Credentials


#Test Env
#Invoke-DatabaseQuery –query "select * from [dbo].[v_OutlookData]"–isSQLServer –connectionString "Data Source=192.168.0.9;Initial Catalog=eHR;User Id=sa;Password=Pass.2016"

#Prod Env
Invoke-DatabaseQuery –query "select * from [dbo].[v_OutlookData]" –isSQLServer –connectionString "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"


Remove-PSSession $exchangeSession

