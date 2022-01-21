$source_dir = "C:\Users\Msr\Desktop\demo\"

## Checking if file exist for todays date
$File = Get-ChildItem -Path $source_dir -Filter 'test*' | Where-Object {$_.LastWriteTime -gt (Get-Date).Date}
$FullFilePath = $source_dir + $File.Name$lastWrite = (get-item $FullFilePath).LastWriteTime

$timespan = new-timespan -days 0 -hours 0 -minutes 5

if(Test-Path -Path $FullFilePath -PathType Leaf) {
	if (((get-date) - $lastWrite) -gt $timespan) {
		echo "file found but its not genrated in last 5 mins, skipping file, exiting program."
	} 
	else {
		echo "new file"
	}
}

else {
	echo "file not found"
}

