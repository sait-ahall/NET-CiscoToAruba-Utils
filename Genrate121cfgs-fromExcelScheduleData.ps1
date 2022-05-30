#use the import excel module to read in the sheet
try {
    $exData = Import-Excel 'C:\Users\ahall\Southern Alberta Institute of Technology\Campus Network Upgrade Project - General\OCM\Schedule\Burns Switch Upgrade schedule.xlsx'
}
catch {
    write-host "File was locked"
    break
}


#Create a new object to finessse the switch and room data

$objOut = @()

# set the order value to generate the specific switch configs:
# make sure that the switch configs are backed up on the switches and are downloaed to Y:
$exdataSelected = $exdata #| Where-Object order -in (8)
foreach ($row in $exdataSelected) {

    if ($row.DeviceName) {
        $x = New-Object -TypeName psobject -Property @{
            RoomID         = $row.RoomID
            ConfigFileName = 'cisco' + $row.DeviceName + '-confg-dump'
            DeviceName     = 'cisco' + $row.DeviceName
            Order          = [int]$row.order
            UGWeek         = $row.'Planned upgrade week'
            sshCommand     = 'ssh ' + 'cisco' + $row.DeviceName
            BackupCommand  = 'copy running-config tftp://10.194.133.10/cisco' + $row.DeviceName + '-confg-dump'
        }

        $objOut += $x
    }
}


$objOutSorted = ($objOut | Sort-Object order)

# create command file for cut and paste
<# "" | Out-File -Path 'c:\foo\backupcmds.txt'

foreach ($row in $objOutSorted) {
    $row.UGWeek | Out-File -Path 'c:\foo\backupcmds.txt' -Append
    $row.sshCommand | Out-File -Path 'c:\foo\backupcmds.txt' -Append
    #"wr" | Out-File -Path 'c:\foo\backupcmds.txt' -Append
    $row.BackupCommand | Out-File -Path 'c:\foo\backupcmds.txt' -Append
    "exit`n`n" | Out-File -Path 'c:\foo\backupcmds.txt' -Append

} #>

#create input csv for roomlist, set vars such as output folder for the Convert-CiscoToAruba module
$objOutRooms = @()

$outputBaseDir = 'c:\foo\__out' #test dir
#$outputBaseDir = 'C:\Users\ahall\Southern Alberta Institute of Technology\Campus Network Upgrade Project - General\Deployment Plans'

foreach ($row in $objOutSorted) {
    $x = New-Object -TypeName psobject -Property @{
        CiscoConfigFilePath = "y:\backups\"+ $row.RoomID +'\'+ $row.ConfigFileName
        OutputDir = $outputBaseDir+'\'+ $row.RoomID
        ConversionMethod = 'OneToOne'
        ConsoleVisualizer ='False'
    }
    $objOutRooms += $x
}


$Pline = @()
foreach ($row in $objOutRooms) {

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

# run converion
$pline | Convert-CiscoToAruba -Verbose




