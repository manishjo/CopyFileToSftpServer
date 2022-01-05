## Defining variables used for this program
$source_dir = "\\vmlkrp1nfsm\interfaces\SFTP\"
$ArchivePath ="\\vmlkrp1nfsm\interfaces\SFTP\ARCHIVE"
$PSEmailServer = 'mail-na.enterprise.cmgi.com'

## Load WinSCP .NET assembly 
Add-Type -Path "C:\PatientPointSFTPScript\WinSCPnet.dll"

## Set up SFTP session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = "SFTP.MODUSLINK.COM"
    UserName = "ML_PatientPoint"
    Password = "c9yXbeLyB4SrW5W4"
    SshHostKeyFingerprint = "ssh-rsa 4096 a3:0a:44:65:07:57:59:b9:c8:32:95:14:91:6b:67:a4"
}

## Checking if RMA file exist for todays date
$MailBody = "Time in UTC = "+$(Get-Date -format 'u') + "`n`nChecking if RMA File exists for today's date.";
$File = Get-ChildItem -Path $source_dir -Filter 'RMA_Serialization_Report*' | Where-Object {$_.LastWriteTime -gt (Get-Date).Date}
$FullFilePath = $source_dir + $File.Name


##If RMA file found for todays date copy to sftp server, else inform file does not exist.
if(Test-Path -Path $FullFilePath -PathType Leaf) {
$MailBody += "`nFile found with name = " + $FullFilePath +"`n`nNow i am copying it to SFTP server"

################### Copy file to SFTP Folder ####################

$session = New-Object WinSCP.Session
try{
    $session.Open($sessionOptions)
    $session.PutFiles($FullFilePath, "/Production/").Check()
}
finally{
    $session.Dispose()
}
$MailBody += "`n**RMA File has been copied to SFTP server fine**"

#################### Move file to Archive ######################

$MailBody += "`n`nNow i am moving RMA file to Archive folder."
Get-ChildItem –Path $FullFilePath | Move-Item -Destination $ArchivePath
$MailBody += "`n**RMA File has been moved to Archive folder successfully**"

#################### Send Mail notification #####################

Send-MailMessage -From 'PatientPoint SFTP Script<no_mail@moduslink.com>' -To '<MLBasisSecurity@moduslink.com>' -Cc '<Ravi_Kiran@moduslink.com>','<gita_karle@moduslink.com>','<alejandro_flores@moduslink.com>','<sandra_luxton@moduslink.com>' -Subject 'PatientPoint SFTP Script Run Summary' -Body $MailBody
}

else {
$MailBody +="`n`n**RMA File does not exists for todays date, Program is exiting**"
Send-MailMessage -From 'PatientPoint SFTP Script<no_mail@moduslink.com>' -To '<MLBasisSecurity@moduslink.com>' -Cc '<Ravi_Kiran@moduslink.com>','<gita_karle@moduslink.com>','<alejandro_flores@moduslink.com>','<sandra_luxton@moduslink.com>' -Subject 'PatientPoint SFTP Script Run Summary' -Body $MailBody
}