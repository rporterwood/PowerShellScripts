<#
.Synopsis
   Loops through SFTP server to download files
.DESCRIPTION
   Connects to an SFTP server and uses a known folder list to loop through and download files.
.EXAMPLE
   Just Run the script
.NOTES
   Redacted for public use

#>
#where are we downloading to
$destinationpath = ""
#where is the user list located
$userlist = ""
#get the sftp folder list (users)
$sftpusers = get-content $userlist
#set username for logging in
$sFTPUserName = ""
# Set the IP of the SFTP server
$SftpIp = ''

#datatable for later
$Datatable = New-Object System.Data.DataTable
$Datatable.Columns.Add("User")
$Datatable.Columns.Add("FileName")




#enforce tls1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function GetSudoPassword
{
	#pulls password from encrypted key file
	#code for keyfile creation, if you use this a new keyfile will be generated
	#use accordingly
	
	#$KeyFile = "C:\temp\sftpAES.key"
	#$Key = New-Object Byte[] 16   # You can use 16, 24, or 32 for AES
	#[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
	#$Key | out-file $KeyFile
	
	
	#Code to generate password file in the event sudo password is changed
	
	#$PasswordFile = "C:\temp\sftpPassword.txt"
	#$Password = Read-Host -AsSecureString
	#$Password | ConvertFrom-SecureString -key $Key | Out-File $PasswordFile
	
	
	#just a dummy username for the PSCredential object
	$username = "dummy"
	#path to NTFS controlled key file
	$KeyFile = ""
	$Key = Get-Content $KeyFile
	#path to NTFS controlled password file
	$pwdTxt = Get-Content ""
	$securePwd = $pwdTxt | ConvertTo-SecureString -key $Key
	$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd
	$credObject.GetNetworkCredential().password
	
}

function GetPlink
{
	

	$PlinkLocation = "C:\Program Files\PuTTY" + "\Plink.exe"
	If (-not (Test-Path $PlinkLocation))
	{
		Write-Host "Plink.exe not found, trying to download..."
		mkdir "C:\Program Files\PuTTY" -ErrorAction Continue
		$WC = new-object net.webclient
		$WC.DownloadFile("https://the.earth.li/~sgtatham/putty/latest/w64/plink.exe", $PlinkLocation)
		If (-not (Test-Path $PlinkLocation))
		{
			Write-Host "Unable to download plink.exe, please download from the following URL and add it to the same folder as this script: http://the.earth.li/~sgtatham/putty/latest/x86/plink.exe"
			#Exit
		}
		Else
		{
			$PlinkEXE = Get-ChildItem $PlinkLocation
			If ($PlinkEXE.Length -gt 0)
			{
				Write-Host "Plink.exe downloaded, continuing script"
			}
			Else
			{
				Write-Host "Unable to download plink.exe, please download from the following URL and add it to the same folder as this script: http://the.earth.li/~sgtatham/putty/latest/x86/plink.exe"
				#Exit
			}
		}
	}
}


    #gets password
	$pass = getsudopassword
	#checks for plink and installs it
	GetPlink





foreach ($suser in $sftpusers)
{

		# Set the credentials
		$Password = ConvertTo-SecureString $pass -AsPlainText -Force
		$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $sFTPUserName, $Password
		
		# Establish the SFTP connection
		
		$ThisSession = New-SFTPSession -ComputerName $SftpIp -Credential $Credential -AcceptKey -ErrorAction Stop
        $FilePath = Get-SFTPChildItem -SessionId $ThisSession.SessionId -Path /sftp/$suser/incoming

        New-Item -ItemType Directory -Force -Path ""

        ForEach ($LocalFile in $FilePath)
{
    
#Ignores '.' (current directory) and '..' (parent directory) to only look at files within the current directory
    if($LocalFile.name -eq "." -Or $LocalFile.name -eq ".." )
    {
          #Write-Host "Files Ignored!"
    }
    else
    {
        #Write-Host $LocalFile
        if($LocalFile.LastWriteTime -ge (get-date).AddDays(-1))
        {
        Get-SFTPFile -SessionId $ThisSession.SessionID -LocalPath "" -RemoteFile $localfile.fullname -Overwrite
        #$filelist += $suser + ',' + $localfile.Name
        $row = $Datatable.NewRow()
        $row.User = $suser
        $row.FileName = $localfile.Name

        $Datatable.Rows.Add($row)
        }
        

        #Remove-SFTPItem -SessionId $ThisSession.SessionID -Path $localfile.fullname -Force
    }   
    

}

}

#Disconnect all SFTP Sessions
	Get-SFTPSession | % { Remove-SFTPSession -SessionId ($_.SessionId) }

$curTime = get-date
$date = (get-date).AddDays(-1).ToString("MM-dd-yyyy")

$htmltable = "<table><tr><td>User</td><td>FileName</td></tr>"

if ($Datatable.Rows.Count -ge 1)
{
foreach ($nrow in $Datatable.Rows)
{ 
    $htmltable += "<tr><td>" + $nrow[0] + "</td><td>" + $nrow[1] + "</td></tr>"
}
$htmltable += "</table>"

$body = "The following files were downloaded at $CurTime  to <redacted> <br><br> $htmltable"


Send-MailMessage -From ""  -To "" -Cc "" -Subject "SFTP Files Downloaded $curTime" -Body $body  -SmtpServer "" -BodyAsHtml

}
else
{

Send-MailMessage -From ""  -To "" -Cc "pwood@acbcoop.com" -Subject "No SFTP Files Downloaded $curTime" -Body "There were no files created in the last 24 hours available for download from the FTP server."  -SmtpServer "" -BodyAsHtml

}