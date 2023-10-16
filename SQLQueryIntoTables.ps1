<#
.Synopsis
   Pulls SQL Data and manipulates into HTML table
.DESCRIPTION
   REDACTED
.EXAMPLE
   Just run script
.NOTES
	Info redacted for public viewing
#>
$dt=get-date -Format "MM-dd-yyyy"
Start-Transcript -Path "c:\scripts\gmbump\transcript-$dt.txt"

#Do all the SQL stuff to get data
$database = ""
$servername = ""



$a10 = 0
$a14 = 0



$queryB3 = ""
try {
$b3 = invoke-sqlcmd  -query "$queryB3" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand Submitted
}
catch
{
send-mailmessage -from pwood@acbooop.com -to pwood@acbcoop.com -Subject "error in gm bump" -body "$error" -SmtpServer mail.acbcoop.net
}

$queryC3 = ""

$c3 = invoke-sqlcmd  -query "$queryC3" -ServerInstance "$servername" -ErrorAction 'Stop'| select -expand column1

$queryc4 = ""

$c4 = invoke-sqlcmd  -query "$queryc4" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

$queryC5 = ""

$c5 = invoke-sqlcmd  -query "$queryC5" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

$queryd3 = ""

$d3 = invoke-sqlcmd  -query "$queryd3" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

$queryd4 = ""

$d4 = invoke-sqlcmd  -query "$queryd4" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

$queryd5 = ""

$d5 = invoke-sqlcmd  -query "$queryd5" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

$querye3 = ""

$e3 = invoke-sqlcmd  -query "$querye3" -ServerInstance "$servername" -ErrorAction 'Stop'  | select -expand submitted

#Selected Approved - Cell F3
$queryf3 = ""

$f3 = invoke-sqlcmd -query "$queryf3" -ServerInstance "$servername" -ErrorAction 'Stop'  | select -expand column1

#Selected Denied - Cell F4
$queryf4 = ""

$f4 = invoke-sqlcmd -query "$queryf4" -ServerInstance "$servername" -ErrorAction 'Stop'  | select -expand column1

#Selected Info Requested - Cell F5
$queryf5 = ""

$f5 = invoke-sqlcmd -query "$queryf5" -ServerInstance "$servername" -ErrorAction 'Stop'  | select -expand column1

#Selected Amount Approved - Cell G3
$queryg3 = "U"

$g3 = invoke-sqlcmd -query "$queryg3" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

#Selected Amount Denied - Cell G4
$queryg4 = ""

$g4 = invoke-sqlcmd -query "$queryg4" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

#Selected Amount Info Requested - Cell G5
$queryg5 = ""

$g5 = invoke-sqlcmd -query "$queryg5" -ServerInstance "$servername" -ErrorAction 'Stop' | select -expand column1

$querya10 = ""
$ra10 = invoke-sqlcmd -query "$querya10" -ServerInstance "$servername" -ErrorAction 'Stop'

foreach ($l in $ra10)
{ 
$a10 += $l.count
}


$queryA14 = ""
$ra14 = invoke-sqlcmd -query "$querya14" -ServerInstance "$servername" -ErrorAction 'Stop'

foreach ($l in $ra14)
{ 
$a14 += $l.count
}

$jan1stcurrentyear = (Get-Date).Date.AddDays(-((Get-Date).DayOfYear - 1)).ToString('yyyy-MM-dd')


$AuditcodeQuery = ""

$auditcoderesults = invoke-sqlcmd -query "$AuditcodeQuery" -ServerInstance "$servername" -ErrorAction 'Stop'

$dollars = [decimal]$D3
$nD3 = '{0:C}' -f $dollars
$dollars = [decimal]$D4
$nD4 = '{0:C}' -f $dollars
$dollars = [decimal]$D5
$nD5 = '{0:C}' -f $dollars
$dollars = [decimal]$G3
$nG3 = '{0:C}' -f $dollars
$dollars = [decimal]$G4
$nG4 = '{0:C}' -f $dollars
$dollars = [decimal]$G5
$nG5 = '{0:C}' -f $dollars


$B3E3 = $B3+$E3
$B3B4B5 = $B3+$B4+$B5
$C3C4C5 = $C3+$C4+$C5
$D3D4D5 = $D3+$D4+$D5
$dollars = [decimal]$D3D4D5
$D3D4D5 = '{0:C}' -f $dollars
$E3E4E5 = $E3+$E4+$E5
$F3F4F5 = $F3+$F4+$F5
$G3G4G5 = $G3+$G4+$G5
$dollars = [decimal]$G3G4G5
$G3G4G5 = '{0:C}' -f $dollars
$B3E3 = $B3+$E3
$C3C4C5F3F4F5 = $C3+$C4+$C5+$F3+$F4+$F5
$D3D4D5G3G4G5  = $D3+$D4+$D5+$G3+$G4+$G5
$dollars = [decimal]$D3D4D5G3G4G5
$D3D4D5G3G4G5 = '{0:C}' -f $dollars
$C3F3	= $C3+$F3
$D3G3	=  $D3+$G3
$dollars = [decimal]$D3G3
$D3G3 = '{0:C}' -f $dollars
$C4F4	= $C4+$F4
$D4G4	= $D4+$G4
$dollars = [decimal]$D4G4
$D4G4 = '{0:C}' -f $dollars
$C5F5	= $C5+$F5
$D5G5 = $D5+$G5
$dollars = [decimal]$D5G5
$D5G5 = '{0:C}' -f $dollars

$table2 = 
"<table cellspacing=""0"" style=""border-collapse:collapse; width:956px"">
<tbody>
<tr>
<td style=""background-color:#002060; border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:1px solid black; height:20px; text-align:center; vertical-align:bottom; white-space:nowrap; width:76px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Audit Code</span></span></span></td>
<td style=""background-color:#002060; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:1px solid black; vertical-align:bottom; white-space:nowrap; width:679px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Audit Desc</span></span></span></td>
<td style=""background-color:#002060; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:1px solid black; vertical-align:bottom; white-space:nowrap; width:157px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Group Description</span></span></span></td>
<td style=""background-color:#002060; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:1px solid black; text-align:center; vertical-align:bottom; white-space:nowrap; width:44px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Count</span></span></span></td>
</tr>
"



foreach ($result in $auditcoderesults) {
    $value1 = $result.'Audit Code'
    $value2 = $result.'Audit Desc'
    $value3 = $result.'Group Description'
    $value4 = $result.'Count'

$table2+=
"<tr>
<td style=""border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; height:20px; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$value1</span></span></span></td>
<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$value2</span></span></span></td>
<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$value3</span></span></span></td>
<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$value4</span></span></span></td>
</tr>"

}

$table2 += 
"</tbody>
</table>
"


$table1 = 
"
<table cellspacing=""0"" style=""border-collapse:collapse; width:1141px"">
	<tbody>
		<tr>
			<td style=""background-color:#002060; border-bottom:1px solid black; border-left:1px solid black; border-right:none; border-top:1px solid black; height:21px; vertical-align:bottom; white-space:normal; width:125px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Cummulative Stats&nbsp;</span></span></span></td>
			<td colspan=""3"" style=""background-color:#002060; border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:1px solid black; text-align:center; vertical-align:bottom; white-space:nowrap; width:248px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Single Claims</span></span></span></td>
			<td colspan=""3"" style=""background-color:#002060; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:1px solid black; text-align:center; vertical-align:bottom; white-space:nowrap; width:248px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Multiple Claims</span></span></span></td>
			<td colspan=""3"" style=""background-color:#002060; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:1px solid black; text-align:center; vertical-align:bottom; white-space:nowrap; width:248px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif"">Total Claims</span></span></span></td>
		</tr>
		<tr>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; height:20px; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Status</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Submitted</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Selected</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Selected Amount</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Submitted</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Selected</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Selected Amount</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Submitted</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Selected</span></span></span></td>
			<td style=""background-color:#d9d9d9; border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Selected Amount</span></span></span></td>
		</tr>
		<tr>
			<td style=""border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; height:20px; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Approved</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$B3</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C3</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$nD3&nbsp;</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$E3</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$F3</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$nG3&nbsp;</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$B3E3</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C3F3</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$D3G3&nbsp;</span></span></span></td>
		</tr>
		<tr>
			<td style=""border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; height:20px; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Denied</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">N/A</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C4</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$nD4&nbsp;</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">N/A</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$F4</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$nG4&nbsp;</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">N/A</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C4F4</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$D4G4&nbsp;</span></span></span></td>
		</tr>
		<tr>
			<td style=""border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; height:20px; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Info Requested</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">N/A</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C5</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$nD5&nbsp;</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">N/A</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$F5</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$nG5&nbsp;</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">N/A</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C5F5</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$D5G5&nbsp;</span></span></span></td>
		</tr>
		<tr>
			<td style=""border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; height:20px; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">Total</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$B3B4B5</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C3C4C5</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$D3D4D5&nbsp;</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$E3E4E5</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$F3F4F5</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$G3G4G5&nbsp;</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$B3E3</span></span></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$C3C4C5F3F4F5</span></span></strong></span></td>
			<td style=""border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><strong><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">&nbsp;$D3D4D5G3G4G5&nbsp;</span></span></strong></span></td>
		</tr>
	</tbody>
</table>

<p>&nbsp;</p>

<table cellspacing=""0"" style=""border-collapse:collapse; width:197px"">
	<tbody>
		<tr>
			<td colspan=""2"" rowspan=""2"" style=""background-color:#002060; border-bottom:.7px solid black; border-left:1px solid black; border-right:.7px solid black; border-top:1px solid black; height:40px; text-align:center; vertical-align:bottom; white-space:normal; width:197px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif""># of Unique BACs that submitted claims since 1/1/23</span></span></span></td>
		</tr>
		<tr>
		</tr>
		<tr>
			<td colspan=""2"" style=""border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:1px solid black; height:20px; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$A10</span></span></span></td>
		</tr>
		<tr>
			<td style=""border-bottom:none; border-left:none; border-right:none; border-top:none; height:20px; vertical-align:bottom; white-space:nowrap"">&nbsp;</td>
			<td style=""border-bottom:none; border-left:none; border-right:none; border-top:none; vertical-align:bottom; white-space:nowrap"">&nbsp;</td>
		</tr>
		<tr>
			<td colspan=""2"" rowspan=""2"" style=""background-color:#002060; border-bottom:.7px solid black; border-left:1px solid black; border-right:.7px solid black; border-top:1px solid black; height:40px; text-align:center; vertical-align:bottom; white-space:normal; width:197px""><span style=""font-size:15px""><span style=""color:white""><span style=""font-family:Calibri,sans-serif""># of Unique BACs<br />
			selected since 1/1/23</span></span></span></td>
		</tr>
		<tr>
		</tr>
		<tr>
			<td colspan=""2"" style=""border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:1px solid black; height:20px; text-align:center; vertical-align:bottom; white-space:nowrap""><span style=""font-size:15px""><span style=""color:black""><span style=""font-family:Calibri,sans-serif"">$A14</span></span></span></td>
		</tr>
	</tbody>
</table>

<p>&nbsp;</p>

<p>&nbsp;</p>

$table2

"



$date = (Get-Date).ToString('MMddyy')


send-mailmessage -from "" -to "" -Subject "" -BodyAsHtml $table1 -SmtpServer ""


stop-transcript


