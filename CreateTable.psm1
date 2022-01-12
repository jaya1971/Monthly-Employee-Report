function NewUsrTable {
    Param(
        [Parameter (Mandatory = $true)][string] $TableName,
        [Parameter (Mandatory = $true)][object[]] $usrArray,
        [Parameter (Mandatory = $false)][string] $Column1,
        [Parameter (Mandatory = $false)][string] $Column2,
        [Parameter (Mandatory = $false)][string] $Column3,
        [Parameter (Mandatory = $false)][string] $Column4,
        [Parameter (Mandatory = $false)][string] $Column5

    )
  Process{    
    [datetime]$DateTest = get-date -Format d
    # Table name
    $tabName = $TableName

    # Create Table object
    $table = New-Object system.Data.DataTable $tabName

    # Define Columns
    $TColumn1 = New-Object system.Data.DataColumn $Column1,([string])
    $TColumn2 = New-Object system.Data.DataColumn $Column2,([string])
    $TColumn3 = New-Object system.Data.DataColumn $Column3,([string])
    $TColumn4 = New-Object system.Data.DataColumn $Column4,([string])
    $TColumn5 = New-Object system.Data.DataColumn $Column5,([string])

    # Add the Columns
    $table.columns.add($TColumn1)
    $table.columns.add($TColumn2)
    $table.columns.add($TColumn3)
    $table.columns.add($TColumn4)
    $table.Columns.Add($TColumn5)
    write-host $table
    ForEach ($user in $usrArray ) 
    {
        if($user.HireDate){$usrDate = $user.HireDate}
        if($user.BirthDate){$usrDate = $user.BirthDate}
        $Years = New-TimeSpan -Start $usrDate -End $DateTest
        $FirmYears = [math]::Floor($years.Days/365)
        # Create a row
        $row = $table.NewRow()

        # Enter data in the row
        $row.$TColumn1 = ($user.DisplayName)
        $row.$TColumn2 = ($user.mail)
        $row.$TColumn3 = ($user.physicalDeliveryOfficeName)
        $row.$TColumn4 = ($usrDate)
        $row.$TColumn5 = ($FirmYears)
 
        # Add the row to the table
        $table.Rows.Add($row)
    }
    return $table}
}

function TitleTable {
    Param(
        [Parameter (Mandatory = $true)][string] $TableName,
        [Parameter (Mandatory = $true)][object[]] $usrArray,
        [Parameter (Mandatory = $false)][string] $Column1,
        [Parameter (Mandatory = $false)][string] $Column2,
        [Parameter (Mandatory = $false)][string] $Column3,
        [Parameter (Mandatory = $false)][string] $Column4,
        [Parameter (Mandatory = $false)][string] $Column5

    )
    #write-host $usrArray
  Process{    
    # Table name
    $tabName = $TableName

    # Create Table object
    $table = New-Object system.Data.DataTable $tabName

    # Define Columns
    $TColumn1 = New-Object system.Data.DataColumn $Column1,([string])
    $TColumn2 = New-Object system.Data.DataColumn $Column2,([string])
    $TColumn3 = New-Object system.Data.DataColumn $Column3,([string])
    $TColumn4 = New-Object system.Data.DataColumn $Column4,([string])
    $TColumn5 = New-Object system.Data.DataColumn $Column5,([string])

    # Add the Columns
    $table.columns.add($TColumn1)
    $table.columns.add($TColumn2)
    $table.columns.add($TColumn3)
    $table.columns.add($TColumn4)
    $table.Columns.Add($TColumn5)
    write-host $table
    ForEach ($user in $usrArray ) 
    {
        # Create a row
        $row = $table.NewRow()

        # Enter data in the row
        $row.$TColumn1 = ($user.Name)
        $row.$TColumn2 = ($user.Email)
        $row.$TColumn3 = ($user.Office)
        $row.$TColumn4 = ($user.Title)
        $row.$TColumn5 = ($user.'Title Date')
 
        # Add the row to the table
        $table.Rows.Add($row)
    }
    return $table}
}
