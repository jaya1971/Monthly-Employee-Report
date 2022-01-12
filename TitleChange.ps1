add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
Import-Module ActiveDirectory
Import-Module SimplySQL
$ThisYear = (get-date).Year
$pass= get-content 'PATHTOPSWD_FILE\psswd_file.txt' | ConvertTo-SecureString
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist "sqlReader", $pass
Open-SqlConnection -Server SQL_SERVER -Database "UserDB" -Credential $Cred -ConnectionName "CONNECTION_NAME"

#This info is originally pulled from AD but this table is populated each night with active users and is faster to pull from.
$CurrentADUsers = Invoke-SqlQuery "select * from AD_ActiveUsers" -ConnectionName "CONNECTION_NAME"

Invoke-SqlUpdate "drop table if exists AD_Title_Change_temp" -ConnectionName "CONNECTION_NAME"
Invoke-SqlUpdate "Create table AD_Title_Change_Temp (Name varChar(512),Email varChar(512),Office varChar(512),Title varChar(512),[Title Date] DATE)" -ConnectionName "CONNECTION_NAME"
 
foreach($user in $CurrentADUsers){
    $ADEmail = $user.email
    $ADName = $user.Name
    $ADOffice = $user.'Physical Location'
    $ADSama = $user.'Logon Name (pre-Windows 2000)'

    #grabs all title changes that occurred this year and populates a table that will be pulled from when running employee report.
    if(Get-ADReplicationAttributeMetadata -Object (Get-Aduser $ADSama) -Server DC1 | Where-Object {($_.AttributeName -eq 'title') -and ($_.LastOriginatingChangeTime.Year -eq $ThisYear)})
    {
        $titleUser = Get-ADReplicationAttributeMetadata -Object (Get-Aduser $ADSama) -Server DC1 | Where-Object {($_.AttributeName -eq 'title') -and ($_.LastOriginatingChangeTime.Year -eq $ThisYear)} `
        | Select-Object @{n='Name';e={$ADName}}, @{n='Email';e={$ADEmail}},AttributeName,AttributeValue,LastOriginatingChangeTime,LastOriginatingChangeDirectoryServerInvocationId
        $Title = $titleUser.AttributeValue
        $TitleDate = $titleUser.LastOriginatingChangeTime

        Invoke-SqlUpdate "Insert into AD_Title_Change_Temp (Name,Email,Office,Title,[Title Date]) VALUES ('$($ADName -replace "'","''")','$ADEmail','$ADOffice','$Title','$TitleDate')" -ConnectionName "CONNECTION_NAME"
    }
}


$SQLSwap = "BEGIN TRANSACTION 
IF OBJECT_ID('AD_Title_Change_backup', 'U') IS NOT NULL
BEGIN
drop table if exists AD_Title_Change_backup
END
IF OBJECT_ID('AD_Title_Change_temp', 'U') IS NOT NULL
BEGIN
EXEC sp_rename 'AD_Title_Change','AD_Title_Change_backup';
WAITFOR DELAY '00:00:05.000'
EXEC sp_rename 'AD_Title_Change_temp','AD_Title_Change';
END
COMMIT TRANSACTION"

Invoke-SqlUpdate $SQLSwap -ConnectionName "CONNECTION_NAME"
