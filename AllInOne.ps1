
#===============SANDVIK.VBS============

$path_server_names = "C:\DBA\test\Servers.txt"
if(Test-Path $path_server_names){
    Write-Host -ForegroundColor Green "File path is valid."
    $servers_list = Get-Content ($path_server_names)
    Write-Host -ForegroundColor Green "Server list loaded."
}else{
    Write-Warning ("Please create a file named $path_server_names containing a list of computers to run script on, one per line. Press any key to exit..")
    start-sleep 5
    return
}
$path_raw_data = "C:\DBA\test\Raw_data.csv"
$db_name = "master"
foreach($server in $servers_list){
    Write-Host -ForegroundColor Green "Running query for $server"
    Set-Content $path_raw_data ""
    Invoke-Sqlcmd -ServerInstance $server -Database $db_name -InputFile "C:\DBA\test\sql_query.sql" | Export-Csv $path_raw_data -Delimiter "," -NoTypeInformation
    $cur_date = Get-Date -format "yyyy-MM-dd hh:mm:ss"
    $conv = Get-Content C:\dba\test\Raw_data.csv
    $first = $conv[0]
    $conv | Where-Object { $_ -ne $first } | out-file C:\dba\test\Raw_data.csv
    foreach($line in Get-Content $path_raw_data){
        $slq_insert_query_first = "INSERT INTO BackupReportDump ( Server_Name, Physicalname, database_id, dbname, status, datafiles, data_mb, logfiles, log_mb, recovery_model, compatibilityLevel, creationdate, page_verify_option, fullbackup, Diffbackup, Transaction_log, date)  VALUES ('" 
        $slq_insert_query_end = "');"
        $objs = $line.Split(",")
        $sql_insert_query_middle = "$server','"
        forEach($obj in $objs){
            $sql_insert_query_middle += $obj + "','"
        }
        $sql_insert_query_middle+= $cur_date    
        $final_insert_query = $slq_insert_query_first + $sql_insert_query_middle + $slq_insert_query_end
        $final_insert_query = $final_insert_query -replace '"',""
        Set-Content "C:\DBA\test\sql_insert_query.sql" $final_insert_query
        $final_insert_query
        Write-Host -ForegroundColor Green "Insert query prepared"
        Invoke-Sqlcmd -ServerInstance "." -Database "PowerSQL" -InputFile "C:\DBA\test\sql_insert_query.sql"    
    }
}
#================END====================

#=============Daily_Report.vbs==========
Remove-Item C:\dba\test\Daily_Report.csv -Force -ErrorAction Ignore
Remove-Item C:\dba\test\rpt_query.sql -Force -ErrorAction Ignore
foreach($server in $servers_list){
    $rpt_query = "select Server_Name, fullbackup from BackupReportDump where Server_Name='$server' and Date >= '$cur_date'"
    #$rpt_query
    Add-Content C:\dba\test\rpt_query.sql $rpt_query
    $db_name = "PowerSql"
    $qout=Invoke-Sqlcmd -ServerInstance $server -Database $db_name -query $rpt_query 
    $flag=$false
    foreach($q in $qout){
      if(-not($q[1] -eq "Backup Ok")){
          $flag=$true
      }
    }
    if($flag){
        $str="$server, Failed"
        Add-Content C:\dba\test\Daily_Report.csv $str
    }else{
        $str="$server, Success"
        Add-Content C:\dba\test\Daily_Report.csv $str
    }
}

#===============END=====================

#===============Trend.VBS===============


#===============END=====================