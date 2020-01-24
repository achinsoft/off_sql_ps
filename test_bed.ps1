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
foreach($server in $servers_list){
    $rpt_query = "select Server_Name, fullbackup from BackupReportDump where Server_Name='$server' and Date >= '$cur_date'"
    $rpt_query
}