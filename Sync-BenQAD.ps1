#############################################
# Sync-BenQAD.ps1
# Rewrited: 13-Aug-2018
# Updated: 14-Aug-2018
# Billy Zhou
# This script is used to create and update AD user based on HR BenQ DB
#############################################

# function TelephoneNumber used for produce telephone number based on office number and ext number
function TelephoneNumber($officePhone, $ext)
{
	if ($officePhone -eq $null) 
	{
		$telephoneNumber = $null	
	}
	else 
	{
		if ($ext -eq $null) 
		{
			$telephoneNumber = $officePhone
		}
		else
		{
			$telephoneNumber = $officePhone + "," + $ext
		}
	}
	return $telephoneNumber
}
# function HandleNull is used for handle DBNull value
function HandleNull($oldvalue)
{
	if($oldvalue.Equals([DBNull]::Value))
	{
		$value = $null
	}
	elseif($oldvalue -eq '')
	{
		$value =  $null
	}
	else
	{
		$value = $oldvalue
	}
	return $value
}

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

$connectionStringBenq = "Data Source=97VMDBSERVER.CHOADEN.COM;Initial Catalog=eHR;User Id=exchange;Password=exchange"
$connectionStringCadena = "Data Source=ccnsvwhqdwh02.choaden.com;Initial Catalog=BI_HR;User Id=weaden;Password=adenweaden@123"

$queryBenq = "select EMPLOYEEID,FIRSTNAME,LASTNAME,DisplayName,Initials,Email,Dept,Region,JobTitle,OFFICEPHONE,EXT,MobilePhone,ReportToID,'BENQ' as source
	from [dbo].[v_OutlookData]
    where (leavedate='' or LeaveDate is null) and (OutDate='' or OutDate is null)" 
$queryCadena = "select EmployeeID,FirstName,LastName,iif(secondname='' or secondname is null,name,secondname) as displayname,'' as initials,Email,Department as dept,Region,JobTitle,'' as officephone,'' as ext,'' as mobilephone,SupervisorID as reporttoid,'CADENA' as source 
    from dbo.HR_EMPS_VN 
    where EmployeeID is not NULL and EmployeeID <>'' and Email like '%@adenservices.com' and EmployeeStatus = 'Active'"


#############################################
## Prepare Log File						#####
#############################################

$ScriptFolder = Split-Path $MyInvocation.MyCommand.Definition -Parent
$batchNo =  Get-Date -Format 'yyyyMMddHH'
$LogPath = "C:\log\SyncBenqAD\"
$runningLog = $LogPath + "RunningLog.log"
$syncBenqADLog = $LogPath + "SyncBenqAD" + $batchNo + ".log"
$exclusion = Get-Content ($ScriptFolder + "\exclusion\Sync-HRAD.txt")
if (!(Test-Path $LogPath))
{
    mkdir $LogPath
}

########## Connect to BenQ DB ############

$connectionBenq = New-Object -TypeName System.Data.SqlClient.SqlConnection

$connectionBenq.ConnectionString = $connectionStringBenq
$commandBenq = $connectionBenq.CreateCommand()
$commandBenq.CommandText = $queryBenq
$adapterBenq = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $commandBenq
$datasetBenq = New-Object -TypeName System.Data.DataSet
$adapterBenq.Fill($datasetBenq)
$tableBenq=$datasetBenq.Tables[0]

$connectionBenq.Close()

$connectionCadena = New-Object -TypeName System.Data.SqlClient.SqlConnection

$connectionCadena.ConnectionString = $connectionStringCadena
$commandCadena = $connectionCadena.CreateCommand()
$commandCadena.CommandText = $queryCadena
$adapterCadena = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $commandCadena
$datasetCadena = New-Object -TypeName System.Data.DataSet
$adapterCadena.Fill($datasetCadena)
$tableCadena=$datasetCadena.Tables[0]

$connectionCadena.Close()

$allData = $tableBenq.Rows + $tableCadena.Rows
#$allData = $tableCadena.Rows

$count = 1

"" > $runningLog

#$allData | ogv
#$allData | Export-Csv d:\cadena.csv

########### Traverse the table ############
foreach ($item in $allData)
{
    $email = HandleNull($item.Email.Trim())
    if ($exclusion -notcontains $email)
    {
	    $employeeId = HandleNull($item.EMPLOYEEID.Trim())
	    $firstName = HandleNull($item.FIRSTNAME.Trim())
	    $lastName = HandleNull($item.LASTNAME.Trim())
	    $displayName = HandleNull($item.DisplayName.Trim())
	    $initials = HandleNull($item.initials.Trim())
	    $dept = HandleNull($item.Dept.Trim())
	    $region = HandleNull($item.Region.Trim())
	    $jobTitle = HandleNull($item.jobTitle.Trim())
	    $officePhone = HandleNull($item.OFFICEPHONE)
	    $ext = HandleNull($item.EXT)
	    $mobilePhone = HandleNull($item.MobilePhone)
	    $reportToId = HandleNull($item.ReportToID)
        $ou = HandleNull($item.source)
	    $telephoneNumber = TelephoneNumber -officePhone $officePhone -ext $ext
	    $name = $email.substring(0, $email.IndexOf("@")).Trim()

	    # check manager name
	    $managerName = $null
	    if ($reportToId -ne $null)
	    {	
		    $managerName = get-aduser -Filter {EmployeeID -eq $reportToId} -ErrorAction SilentlyContinue | select samAccountName
	    }
	    # filter user by EmployeeID
	    $adAccount = Get-ADUser -Filter {EmployeeID -eq $employeeId} | Select-Object samAccountName,UserPrincipalName
	    if ($adAccount.Count -ge 2)
	    {
		    if (($adAccount | where {$_.enabled -eq $true}).Count -ge 2)
		    {	
			    $employeeId + "`t" + $email + "`thas more than one enabled AD accounts." >> $syncBenqADLog
		    }
		    $sam = $adAccount[0].SamAccountName
	    }
	    else 
	    {
		    $sam = $adAccount.SamAccountName
    	}
	    if ($email -notlike "*@adenservices.com" -and $email -notlike "*@axingservices.com") 
	    {
		    $count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tEmail format error. AD user will not be created or updated."
		    $count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tEmail format error. AD user will not be created or updated." >> $runningLog
		    $employeeId + "`t" + $name + "`t" + $email + "`tEmail format error. AD user will not be created or updated." >> $syncBenqADLog
	    }
	    # user already exits
	    elseif ($sam -ne $null)
	    {
			# update AD user
	        Set-ADUser $sam `
		        -SurName $lastName `
		        -GivenName $firstName `
		        -DisplayName $displayName `
		        -Department $dept `
		        -Office $region `
		        -Title $jobTitle `
		        -OfficePhone $telephoneNumber `
		        -MobilePhone $mobilePhone `
		        -Initials $initials `
		        -userprincipalname $email `
		        -EmailAddress $email `
		        -Manager $managerName `
		        -Replace @{msExchHideFromAddressLists=$False}

	        $count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tADUser updated."
	        $count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tADUser updated." >> $runningLog
	    }
        else
        {
		    $useraccount = Get-ADUser -Properties employeeid -Filter {sAMAccountName  -eq $name}
		    $useraccount2 = Get-ADUser -Properties employeeid -Filter {userPrincipalName  -eq $email}

		    if($useraccount -ne $Null)
		    {
                $currentEmployeeId = $useraccount.employeeid
                if($currentEmployeeId -ne '' -and $currentEmployeeId -ne $null)
                {
			        $count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tEmployeeid conflict. Current employeeid in AD: " + $currentEmployeeId
			        $count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tEmployeeid conflict. Current employeeid in AD: " + $currentEmployeeId
			        $employeeId + "`t" + $name + "`t" + $email + "`tEmployeeid conflict. Current employeeid in AD: " + $currentEmployeeId >> $syncBenqADLog
                }
                else
                {
                    set-aduser $useraccount.SamAccountName -EmployeeID $employeeId
                    $count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tEmployeeid updated."
			        $count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tEmployeeid updated." >> $runningLog
			        $employeeId + "`t" + $name + "`t" + $email + "`tEmployeeid updated." >> $syncBenqADLog

                }
		    }
		    elseif($useraccount2 -ne $Null)
		    {
			    $count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tUPN already exists. Creation failed."
			    $count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tUPN already exists. Creation failed." >> $runningLog
			    $employeeId + "`t" + $name + "`t" + $email + "`tUPN already exists. Creation failed." >> $syncBenqADLog
		    }
		    else
		    {
			    $ouPath = "OU="+ $ou +",OU=ADEN-Users,DC=CHOADEN,DC=COM"
			    if([adsi]::Exists("LDAP://$ouPath"))
			    {
                    $password = "Aden@123" + $employeeId
                    New-ADUser $name `
					    -SamAccountName $name `
					    -userprincipalname $email `
					    -Surname $lastName `
					    -GivenName $firstName `
					    -DisplayName $displayName `
					    -Department $dept `
					    -Office $region `
					    -Title $jobTitle `
					    -OfficePhone $telephoneNumber `
					    -MobilePhone $mobilePhone `
					    -Initials $initials `
					    -EmailAddress $email `
					    -EmployeeID $employeeId `
					    -Manager $managerName `
					    -Path $ouPath   `
					    -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
					    -enabled $true
					    $count.ToString() + $employeeId + "`t" + $name + "`t" + $password +"`t" + $email + "`tAD account created."
					    $count.ToString() + $employeeId + "`t" + $name + "`t" + $password +"`t" + $email + "`tAD account created." >> $runningLog
					    $employeeId + "`t" + $name + "`t" + $email + "`tAD account created." >> $syncBenqADLog
			    }
			    else 
			    {
				    "OU path error. Script will terminate..." >> $syncBenqADLog
				    break
			    }
		    }
        }
    }
    else
    {
        $count.ToString() + "`t" + $email + "`tskipped for in exclusion list."
		$count.ToString() + "`t" + $email + "`tskipped for in exclusion list." >> $runningLog
    }
    $count ++
}
if (Test-Path $syncBenqADLog)
{
    $msg=New-Object System.Net.Mail.MailMessage
    $msg.To.Add("billy.zhou@adenservices.com,global.itsup@adenservices.com")
    $msg.From=New-Object System.Net.Mail.MailAddress("log@adenservices.com");
    $msg.Subject="SyncBenqADLog" + $batchNo
    $msg.Body=Get-Content $syncBenqADLog -Raw
    $client=New-Object System.Net.Mail.SmtpClient("smtprelay.it.adenservices.com")
    $client.Send($msg)
}
Get-PSSession | Remove-PSSession

	