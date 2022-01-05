## Defining variables used for this program
$source_dir = "<Source Dir where file will be genrated>"
$ArchivePath ="<Arhive folder Dir>"
$PSEmailServer = '<smtp mail server>'

## Load WinSCP .NET assembly
Add-Type -Path "<path>\WinSCPnet.dll"

## Set up SFTP session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
Protocol = [WinSCP.Protocol]::Sftp
HostName = "<sftp host name>"
UserName = "<user>"
Password = "<password>"
SshHostKeyFingerprint = "<host key to connect to sftp server>"
}

## Checking if file exist for todays date
$MailBody = "Time in UTC = "+$(Get-Date -format 'u') + "`n`nChecking if RMA File exists for today's date.";
$File = Get-ChildItem -Path $source_dir -Filter '<file name>*' | Where-Object {$_.LastWriteTime -gt (Get-Date).Date}
$FullFilePath = $source_dir + $File.Name


##If file found for todays date copy to sftp server, else inform file does not exist.
if(Test-Path -Path $FullFilePath -PathType Leaf) {
$MailBody += "`nFile found with name = " + $FullFilePath +"`n`nNow i am copying it to SFTP server"

################### Copy file to SFTP Folder ####################

$session = New-Object WinSCP.Session
try{
    $session.Open($sessionOptions)
    $session.PutFiles($FullFilePath, "/<path of sftp server>/").Check()
}
finally{
    $session.Dispose()
}
$MailBody += "`n**File has been copied to SFTP server fine**"

#################### Move file to Archive ######################

$MailBody += "`n`nNow i am moving file to Archive folder."
Get-ChildItem –Path $FullFilePath | Move-Item -Destination $ArchivePath
$MailBody += "`n**RMA File has been moved to Archive folder successfully**"

#################### Send Mail notification #####################

Send-MailMessage -From '<sender mail id>' -To '<reciever mail id>' -Subject 'SFTP Script Run Summary' -Body $MailBody
}

else {
$MailBody +="`n`n**File does not exists for todays date, Program is exiting**"
Send-MailMessage -From '<sender mail id>' -To '<reciever mail id>' -Subject 'SFTP Script Run Summary' -Body $MailBody
}