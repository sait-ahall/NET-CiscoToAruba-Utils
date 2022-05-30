# import the rooms for the pipline


$roomListCSV = 'C:\foo\roomlist.csv'
$RoomList = Import-CSV -path $roomListCSV

$Pline = @()
foreach ($row in $RoomList) {

    #convert the text to a bool for the visualiser flag
    if ($row.ConsoleVisualizer.toupper() -eq 'FALSE') {
        $ShowVis = $false
    } else {
        $ShowVis = $true
    }

    $x = New-Object -TypeName psobject -Property @{
        CiscoConfigFilePath = $row.CiscoConfigFilePath
        OutputDir           = $row.OutputDir
        ConversionMethod    = $row.ConversionMethod
        ConsoleVisualizer   = $ShowVis
    }

    $pline += $x
}

#$pline | Convert-CiscoToAruba
