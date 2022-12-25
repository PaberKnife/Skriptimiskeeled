$tspan = Read-Host "Ajavahemik"
while ($true) {
    $path = Read-Host "Faili salvestuskoht"
    $check = Test-Path $path -ErrorAction Stop
    if ($check) {
        break 
    }
}

if(($path.EndsWith('.txt') -eq $False) -and ($path.EndsWith('.csv') -eq $False) -and ($path.EndsWith('.xml') -eq $False)) {
    $path += '/süsteemi_sündmused.txt'
}

$count = 0
for ($i = 0; $i -lt $tspan; $i++) {
    $startTime = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-$i-1)
    $endTime = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-$i)
    $winEvents = Get-WinEvent @{ logname="System"; level=1,2,3; starttime=$startTime; endtime=$endTime} | Measure-Object
    $count += $winEvents.Count
}

 $lastEvent = get-eventlog system | Sort-Object timegenerated | Select-Object -first $count
 $topEvent = $lastEvent | group-object source, instanceid | sort-object count -descending
 $addVariable = 0
 $topEvent | foreach {$addVariable+=2}
 $collection = [Object[]]::new($addVariable)
 $addVariable = 0
 $topEvent | foreach {($collection[$addVariable] = $_.name.split(",")[0].Trim()) -and ($collection[$addVariable+1] = $_.name.split(",")[1].Trim()) -and ($addVariable+=2)}
 $find = $lastEvent | sort-object instanceid,source

 for ($i = 0; $i -lt $collection.Count-1;$i+=2) {
    "Nimetus: " + $collection[$i] + "`t ID: " + $collection[$i+1] >> $path
    $find |Sort-Object timegenerated | Where-Object{($_.source -eq $collection[$i]) -and ($_.instanceid -eq $collection[$i+1])} | Format-Table timegenerated, message >> $path
}

$output = "Faili salvestuskoht: " + $path
write-output $output