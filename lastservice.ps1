<#
    .Synopsis
        Gets a distribution list, runs a query based on the current user in the list, and then e-mails them
        a csv file of the query results.
    .Examples
#>

#Collecting information for the script to run.
$instance = 'LBSSQL\ecris'
$server = 'LBSSQL'
$db = 'DataLBS'
$path = 'C:\psout\users.csv'
$cred = get-credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://lbsmail.lbscares.com/PowerShell/ -Authentication Kerberos -Credential $cred

Import-PSSession $Session -AllowClobber
$list = Get-DistributionGroupMember 'housing connect' | select displayname, primarysmtpaddress | export-csv $path

#Imports the list of users and runs the query for each user.
import-csv $path | foreach {

$query = "select pr_name, 
st_fname+' '+st_lname as 'Assigned Staff', 
cl_rcdid as 'PCE#',
cl_caseno as 'Member ID', 
cl_fname, 
cl_lname,
SA_FRMDT as 'Staff Assignment Date', 
(
select Cast(max(SA_SRVDATE)as date)
from DCS_PCHSALPF 
where saf_cltid = CL_RCDID
and saf_stfid = ST_RCDID
and sa_face = 'Y'
) as 'Last Service',

(
select datediff(d, max(sa_srvdate), getdate())
from DCS_PCHSALPF 
where saf_cltid = CL_RCDID
and saf_stfid = ST_RCDID
and sa_face = 'Y'
) as 'Days Since Last Service'

from dcs_pchcltpf 
join dcs_pchprvpf on clf_prvid = PR_RCDID
join dcs_pchstfpf on clf_stfid = ST_RCDID
join dcs_pchadmpf on clf_admid = AD_RCDID
join dcs_pchasapf on adf_asaid = sa_rcdid
where cl_status = 'O'
and st_fname+' '+st_lname = '$($_.displayname)'
ORDER BY [Last Service] ASC"

$path2 = "C:\psout\$($_.displayname).csv"
Invoke-Sqlcmd -hostname $server -ServerInstance $instance -database $db -query $query -OutputAs DataRows | export-csv -path $path2

#Creating the variables to email the user.
$Smtp ="outlook.office365.com"
$To = "$($_.primarysmtpaddress)"
$From = "joshm@lbscares.com"
$Subject = "Last service by caseload for $($_.displayname)"
$Body = "The attached spreadsheet has the last service date for each client on your caseload. If a record is blank or Null then they've never seen you."

Send-MailMessage -SmtpServer $smtp -to $to -from $from -subject $subject -body $body -attachments $path2 -credential $cred -UseSsl -port 587
}