<#
    .Synopsis
        Gets a distribution list, runs a query based on the current user in the list, and then e-mails them
        a csv file of the query results.
    .Examples
#>

#Collecting information for the script to run.
$instance = 'lbssql\ecris'
$server = 'LBSSQL'
$db = 'DataLBS'
$path = 'C:\psout\users.csv'
$cred = get-credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://lbsmail.lbscares.com/PowerShell/ -Authentication Kerberos -Credential $cred
$list = Get-DistributionGroupMember csm | select displayname, primarysmtpaddress | export-csv $path

Import-PSSession $Session -AllowClobber

#Imports the list of users and runs the query for each user.
import-csv $path | foreach {

$query = "select 
	st_fname, 
	st_lname, 
	cl_fname, 
	cl_lname, 
	cast(dateadd(year, 1, max(a.dc_docdt)) as date) 'Locus',
	cast(dateadd(year, 1, max(b.dc_docdt)) as date) 'IBPS',
	cast(dateadd(year, 1, max(c.dc_docdt)) as date) 'Arizona',
	cast(dateadd(year, 1, max(d.dc_docdt)) as date) 'Consent',
	cast(dateadd(year, 1, max(e.dc_docdt)) as date) 'Aims',
	cast(dateadd(year, 1, max(i.dc_docdt)) as date)'Med Consent',
	cast(dateadd(year, 1, max(j.dc_docdt)) as date)'Psych Eval',
	cast(dateadd(year, 1, max(k.dc_docdt)) as date)'Treatment Plan'


from dcs_pchcltpf
	left join dcs_pchstfpf on clf_stfid = ST_RCDID
	left join dcs_pchdocpf as a on a.dcf_cltid = cl_rcdid
		and a.DCF_DOCTYP = '16150'
	left join dcs_pchdocpf as b on b.dcf_cltid = cl_rcdid
		and b.dcf_doctyp = '16156'
	left join dcs_pchdocpf as c on c.dcf_cltid = cl_rcdid
		and c.dcf_doctyp = '15001'
	left join dcs_pchdocpf as d on d.dcf_cltid = cl_rcdid
		and d.DCF_DOCTYP = '16070'
	left join dcs_pchdocpf as e on e.dcf_cltid = cl_rcdid
		and e.DCF_DOCTYP = '14974'
	left join dcs_pchdocpf as i on i.dcf_cltid = cl_rcdid
		and i.dcf_doctyp = '15890'
	left join dcs_pchdocpf as j on j.dcf_cltid = cl_rcdid
		and j.dcf_doctyp in ('13981' , '15921')
	left join dcs_pchdocpf as k on k.dcf_cltid = cl_rcdid
		and k.dcf_doctyp = '13973'
where cl_status = 'O'
	and st_fname+' '+st_lname = '$($_.displayname)'
group by st_fname, st_lname, cl_fname, CL_LNAME
order by [Treatment Plan]"

$path2 = "C:\psout\$($_.displayname).csv"
Invoke-Sqlcmd -hostname $server -ServerInstance $instance -database $db -query $query -OutputAs DataRows | export-csv -path $path2

#Creating the variables to email the user.
$Smtp ="outlook.office365.com"
$To = "$($_.primarysmtpaddress)"
$From = "joshm@lbscares.com"
$Subject = "Caseload Report for $($_.displayname)"
$Body = "The attached spreadsheet has the due dates for your entire caseload. If a record is blank or Null then there isn't completed documentation."

Send-MailMessage -SmtpServer $smtp -to $to -from $from -subject $subject -body $body -attachments $path2 -credential $cred -UseSsl -port 587
}