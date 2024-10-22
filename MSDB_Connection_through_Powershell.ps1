$Base_Directory = Split-Path $script:MyInvocation.MyCommand.Path
$configurationFile = "$Base_Directory\configurationFile.cnf"
$currentDate = Get-Date
$currentDate = Get-Date -UFormat "%d %b %Y"

Get-Content $configurationFile | ForEach-Object -Begin {$h=@{}} -Process {$k= [regex]::Split($_, "="); if(($k[0].CompareTo("") -ne 0) `
    -and ($k[0].StartWith("#") -ne $True) -and ($k[0].StartWith("[") -ne $True)) {$h.Add($k[0], $k[1])}}




function writeLog
    {
        Param ([String]$logString)
        $logfile_path = "$Base_Directory\logs.log"
        $DateTime = "{0:yy/MM/dd} {0:HH:mm:ss}" -f (Get-Date)
        $logMessage = "$($DateTime) $($logString)"
        Add-Content $logfile_path -Value $logMessage


    }



function Send-Mail `
    {
        Param(
            [Parameter(Mandatory)]
            [String]$BodyMessage,
            [String]$Subject
        )
        try{
            $From = ($h.Get_Item("From")).split(",")
            $To = ($h.Get_Item("To")).split(",")
            $CC = ($h.Get_Item("CC")).split(",")
            $Subject += " "+ $currentDate
            $SMTP_Server = $h.Get_Item("SMTPServer")

            Send-MailMessage -From $From -To $To -Cc $CC -Subject $Subject -Body $BodyMessage -SmtpServer $SMTP_Server -DeliveryNotificationOption OnSuccess -BodyAsHtml
            
            writeLog -logString "mail send Successfully"
        }
        catch{
            writeLog -logString $_.Exception.Message
        
        }

    }



function Execute-SqlQuery{
    param (
        [String]$SqlQuery,
        [System.Object]$SqlConnection
    )
    try{
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $SqlQuery
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        #creating DataSet
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)

        Return DataSet
            
        }
    catch{
        writeLog -logString $_.Exception.Message
        
        }

}



#Main

try{
    $SQLServer = $h.Get_Item("sqlServer")
    $SQLDBName = $h.Get_Item('sqlDBName')
    $UserId = $h.Get_Item('UserId')
    $Password = $h.Get_Item('Password')

    $ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $UserId; Password = $Password"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $ConnectionString
    $SqlConnection.Open()

    $sqlQuery = "SELECT * FROM SYSDBA_STORAGE.dbo.Error_Table where Status2 is NULL"

    $output = Execute-SqlQuery -SqlQuery $sqlQuery -SqlConnection $SqlConnection
    $output1 = $output.Tables[0] | Where-Object {($_.Error).StartWith("Valide Command")} 
    $output2 = $output1 | ConvertTo-Html
    $BodyMessage = "<p>Dear ALL, <br>Please find the datatable:$($output2)</p>"
    Send-Mail -BodyMessage $BodyMessage -Subject $h.Get_Item("Subject")

}
catch{
    writeLog -logString $_.Exception.Message
}
finally{
    if($SqlConnection.State -eq 'Open'){
        $SqlConnection.Close()
        writeLog -logString "SQl connection closed successfully"
    }

}
