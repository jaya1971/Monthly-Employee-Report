Import-Module -Name ActiveDirectory
Import-Module SimplySQL
Import-Module 'PATHTOPSM\CreateTable.psm1' -Force
#grabbing encrypted pw
[Byte[]] $key = (1..32)
$pass = get-content ".\pswdfile.txt" | ConvertTo-SecureString -Key $key
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist "sqlreader", $pass
Open-SqlConnection -Server SQL_SERVER -Database "UserDB" -Credential $Cred -ConnectionName "CONNECTION_NAME"

#Create different date formats for you queries
[string]$Date = get-date -Format d
$ThisMonth = (Get-Date).Month
$LastMonth = (Get-Date).AddMonths(-1).Month
$NextMonth = (Get-Date).AddMonths(+1).Month
$LastYear = (get-date).AddYears(-1).Year
$ThisYear = (get-date).Year
$subject = "Employee report - " + $date
$smtpServer = "SMTP_Server"
$FromAddress = "Report_Sender@emaildomain.com"

function UsrInfo{
    Param(
        [Parameter (Mandatory = $true)][object[]] $OUArray,
        [Parameter (Mandatory = $true)][string] $Offices,
        [Parameter (Mandatory = $false)][string] $EmailAddress
    )

process{
  $ToAddress = $EmailAddress
    ##AD Queries
    foreach($OU in $OUArray){
        #Calculates Hires for last month from when report is triggered.
        [Array]$NewHires += Get-ADUser -Filter * -Properties DisplayName,mail,physicalDeliveryOfficeName,HireDate -SearchBase $OU | Where-Object { $_.HireDate.Month -eq $LastMonth -and $_.HireDate.Year -eq $ThisYear} `
          | Select-Object DisplayName,mail,physicalDeliveryOfficeName,@{n="Hire Date";e={$_.HireDate.ToUniversalTime()}}

        #Calculates Birthdays for the current month.
        [Array]$Birthdays += Get-ADUser -Filter * -Properties DisplayName,mail,physicalDeliveryOfficeName,BirthDate -SearchBase $OU | Where-Object {$_.BirthDate.Month -eq $ThisMonth}`
          | Select-Object DisplayName,mail,physicalDeliveryOfficeName,@{n="Birth_Date";e={$_.BirthDate.ToUniversalTime()}}

        #Calculates Anniversaries if employee has been with the company for at least one year
        [Array]$Anniversaries += Get-ADUser -Filter * -Properties DisplayName,mail,physicalDeliveryOfficeName,HireDate -SearchBase $OU | Where-Object ({ $_.HireDate.Month -eq $ThisMonth -and  $_.HireDate.Year -le $LastYear}) `
          | Select-Object DisplayName,mail,physicalDeliveryOfficeName,@{n="HireDate";e={$_.HireDate.ToUniversalTime()}}
    }

    #Title changes are kept in a db originally pulled from AD Replication Attribute Metadata
    $TitleQuery = "declare @LastMonth as datetime
    set @LastMonth = dateadd(MONTH, datediff(MONTH, 0, GETDATE())-1, 0)
    declare @ThisMonth as datetime
    set @ThisMonth = dateadd(MONTH, datediff(MONTH, 0, GETDATE()), 0)
    
    select * from AD_Title_Change where Office IN ($offices)and (cast([title date] as date) >= cast(@LastMonth as date) and  cast([title date] as date) < cast(@ThisMonth as date)) order by [Title Date]"
    $TitleChanges = Invoke-SqlQuery $TitleQuery -ConnectionName "CONNECTION_NAME"
    
    if($NewHires){ $table1 = NewUsrTable -TableName "New Hires" -usrArray $NewHires -Column1 "Name" -Column2 "Email" -Column3 "Office" -Column4 "Hire Date"}
    if($TitleChanges){$table2 = TitleTable -TableName "Title Changes" -usrArray $TitleChanges -Column1 "Name" -Column2 "Email" -Column3 "Office" -Column4 "Title" -Column5 "Title Date"}
    if($Birthdays){$table3 = NewUsrTable -TableName “Birthdays” -usrArray $Birthdays -Column1 "Name" -Column2 "Email" -Column3 "Office" -Column4 "Birth Date"}
    if($Anniversaries){$table4 = NewUsrTable -TableName “Anniversaries” -usrArray $Anniversaries -Column1 "Name" -Column2 "Email" -Column3 "Office" -Column4 "Anniversary Date" -Column5 "Years with the company"}
   
    
    ###############
    # Build Email #
    ###############
    # Creating head style
    $Head = @"
      
    <style>
      body {
        font-family: "Arial";
        font-size: 8pt;
        color: #4C607B;
        }
      th, td { 
        border: 1px solid #e57300;
        border-collapse: collapse;
        padding: 5px;
        }
      th {
        font-size: 1.2em;
        text-align: left;
        background-color: #003366;
        color: #ffffff;
        }
      td {
        color: #000000;
        }
      .even { background-color: #ffffff; }
      .odd { background-color: #bfbfbf; }
    </style>
      
"@

[string]$Hirebody = [PSCustomObject]$table1 | Select-Object -Property "Name","Email","Office","Hire Date"| Sort-Object -Property "Name"  | ConvertTo-HTML `
-head $Head -Body "<font color=`"Black`"><h4>New Hire report</h4></font>"
[string]$Titlebody = [PSCustomObject]$table2 | Select-Object -Property "Name","Email","Office","Title","Title Date"| Sort-Object -Property "Name"  | ConvertTo-HTML `
-head $Head -Body "<font color=`"Black`"><h4>Title Changes report</h4></font>"
[string]$Birthdaybody = [PSCustomObject]$table3 | Select-Object -Property "Name","Email","Office","Birth Date" | Sort-Object -Property "Name" | ConvertTo-HTML `
-head $Head -Body "<font color=`"Black`"><h4>Birthday report</h4></font>"
[string]$Anniversarybody = [PSCustomObject]$table4 | Select-Object -Property "Name","Email","Office","Anniversary Date","Years with Firm" | Sort-Object -Property "Name" | ConvertTo-HTML `
-head $Head -Body "<font color=`"Black`"><h4>Anniversary report</h4></font>"

#Combining multiple HTMLbodies
$body = $HireBody + $Titlebody + $Birthdaybody + $Anniversarybody

# Send the report email
Send-MailMessage -To $ToAddress -Subject $subject -BodyAsHtml $body -SmtpServer $smtpServer -From $FromAddress 
clear-variable NewHires
Clear-Variable TitleChanges
Clear-Variable Anniversaries
Clear-Variable Birthdays
    }
}


#----------------------Break down of Managing Directors and their offices - This section requires considerable setup----------------------------------------
$OUs = "OU=Location1_OU, OU=Company,DC=Domain,DC=com", "OU=Location2_OU, OU=Company,DC=Domain,DC=com","OU=Location3_OU, OU=Company,DC=Domain,DC=com"

#This is office abbreviations we use and passed to the Title SQL query above.
$offices = "'OFF1','OFF2','OFF3'"
$EmailTo = "SiteManager1@emaildomain.com"
UsrInfo -OUArray $OUs -Offices $offices -EmailAddress $EmailTo

$OUs = "OU=Location4_OU, OU=Company,DC=Domain,DC=com", "OU=Location5_OU, OU=Company,DC=Domain,DC=com","OU=Location6_OU, OU=Company,DC=Domain,DC=com"

#This is office abbreviations we use and passed to the Title SQL query above.
$offices = "'OFF4','OFF5','OFF6'"
$EmailTo = "SiteManager2@emaildomain.com"
UsrInfo -OUArray $OUs -Offices $offices -EmailAddress $EmailTo
