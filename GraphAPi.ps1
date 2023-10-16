## this runs in our HR software to automate onboarding


##there is some Active directory work, MS graph API, and generic API work


$token = $Context.CloudServices.GetAzureAuthAccessToken("https://graph.microsoft.com")
$token = $token | ConvertTo-SecureString -AsPlainText -Force

Connect-MgGraph -AccessToken $token

start-sleep 1


try {
$password = "%adm-RandomString,8%"
new-mgUser -UserPrincipalName "%sAMAccountName%@acbcoop.com" -mailnickname "%sAMAccountName%" -DisplayName "%fullname%" -GivenName "%firstname%" -SurName "%lastname%" -UsageLocation "US" -Country "US" -City "Memphis" -State "TN" -PasswordProfile @{"forceChangePasswordNextSignIn" = $true; "password" = "$password"} -AccountEnabled

Set-MgUserLicense -UserId "%sAMAccountName%@acbcoop.com" -Addlicenses @{SkuId = ''} -RemoveLicenses @() 


}
catch {Send-MailMessage -From "" -To "" -Subject "Email not created for %sAMAccountName%" -Body "Please see error: $_.Exception.Message" -SmtpServer "" -port 25
}
try {
Set-AdmAccountPassword -Identity %sAMAccountName% -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
}
catch {$Context.LogMessage("Network Logon password was not set correctly.", "Warning")}
$employeetype = "%employeeType%"
try {
if ($employeetype -eq 'Temp')
{

Set-AdmUser "CN=%name%,OU=,DC=" -Add @{"adm-CustomAttributeText2"=$password} -AdaxesService localhost
}
else
{
    Set-AdmUser "%distinguishedName%" -Add @{"adm-CustomAttributeText2"=$password} -AdaxesService localhost 
}
}
catch
{
Send-MailMessage -From "" -To "" -Subject "User password not set for %sAMAccountName%" -Body "Please see error: $_.Exception.Message" -SmtpServer "" -port 25    
}
start-sleep -Seconds 1
try {
Set-ADUser -Identity %sAMAccountName% -ChangePasswordAtLogon $true
}
catch {$Context.LogMessage("Network Logon password was not set to force a change.", "Warning")}


## this section has to deal with building access controls via API

$Departments = @{

}

try {
$Department = "%department%"
$employeeid = "%employeeID%"

     
$digit1 = $Departments[$Department]
$cardcode = $digit1+$employeeid
start-sleep -seconds 2
Set-AdUser -Identity %sAMAccountName% -EmployeeNumber $cardcode -server newdc2


$body = "Id=0&SaveImage=False&ImageFilename=&KeepImageAspect=False&IsCredentialHolder=True&IsSystemUser=False&ExternalId=$cardcode&PersonalInfo.Id=0&PersonalInfo.FirstName=%firstname%&PersonalInfo.MiddleInitial=&PersonalInfo.LastName=%lastname%&PersonalInfo.DisplayName=&PersonalInfo.EmployeeId=$employeeid&PersonalInfo.Title=&PersonalInfo.Suffix=&PersonalInfo.Department=&ContactInfo.PrimaryEmail=&ContactInfo.SecondaryEmail=&ContactInfo.Company=&ContactInfo.Office=&ContactInfo.Building=&ContactInfo.Position=&ContactInfo.LicensePlateNumber=&ContactInfo.Notes=&ContactInfo.PrimaryPhoneNumber=&ContactInfo.PrimaryExtension=&ContactInfo.CellPhoneNumber=&GroupInfo.AddGroups=1&GroupInfo.RemoveGroups=&CustomFields=&BadgeInfo.TimeZoneBias=0&BadgeInfo.Disabled=False&BadgeInfo.ExpirationDate=&BadgeInfo.ExpirationTime=&BadgeInfo.ActivationDate=$curdate&BadgeInfo.ActivationTime=12:00 AM&BadgeInfo.SiteCode=0&BadgeInfo.CardIssueCode=$cardcode&BadgeInfo.PinCode=&BadgeInfo.Id=0&RoleInfo.Role=None&RoleInfo.Username=&RoleInfo.Password=&RoleInfo.PasswordConfirm=&RoleInfo.PasswordChanged=True"

  
       
invoke-webrequest -method POST -uri http://server:18779/infinias/ia/People?<redacted> -body $body


}
catch
{
    if ($_.Exception.Message -match "The response content cannot be parsed because the Internet Explorer engine is not available, or Internet Explorer's first-launch configuration is not complete. Specify the UseBasicParsing parameter and try again.")
    {
        #do nothing but return this error doesn't seem to affect anything.
        $Context.LogMessage("Internet Explorer Error on Door Setup", "Information")
        return
    }
    else
    {
Send-MailMessage -From "" -To "" -Subject "" -Body "Door Access for %fullname% failed with error $_.Exception.Message" -SmtpServer "" -port 25
    return
}
}





###more MS graph API work


start-sleep -seconds 10
try {
    Connect-MgGraph -Credential $office365Cred

    # Get the service principal for the app you want to assign the user to
    $servicePrincipal = Get-MgServicePrincipal -Filter "displayName eq 'Adobe Identity Management (OIDC)'" | Select-Object -First 1

    $email = "%sAMAccountName%@domain.com"
    $filter = "userPrincipalName eq '$email'"

    # Get new user
    $user = Get-MgUser -Filter $filter | Select-Object displayName, id

    # Add user to Azure App
    # Assuming that the App Role you want is the first one in the list
    New-MgUserAppRoleAssignment -UserId $user.id -AppRoleId $servicePrincipal.AppRoles[0].id -ResourceId $servicePrincipal.id -PrincipalId $user.id
}
catch {
    $Context.LogMessage("Adobe Sync was not enabled.", "Warning")
    Send-MailMessage -From "" -To "" -Subject "Something happened during Adobe AzureAD App adding %sAMAccountName%" -BodyAsHtml "$Error" -SmtpServer "" -Port 25
    $Error.Clear()
}
#distribution lists
$employeetype = "%employeeType%"

try 
{
    
    #Connect-ExchangeOnline -credential $office365Cred
    $Context.CloudServices.ConnectExchangeOnline()
    
if ($employeetype -eq 'Temp')
{
add-distributiongroupmember -Identity "" -Member $email -confirm:$false

}

else
{
add-distributiongroupmember -Identity "" -Member $email -confirm:$false   
 
}
}
catch
{
Send-MailMessage -from "" -to "" -subject "Something happened during email distribution adding %sAMAccountName%" -BodyAsHtml "$Error" -SmtpServer "" -port 25
}
