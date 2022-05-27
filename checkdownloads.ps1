#compares files against the excel file

$rootDir = 'y:\backups'
#import the excel file
$exData = Import-Excel 'C:\Users\ahall\Southern Alberta Institute of Technology\Campus Network Upgrade Project - General\OCM\Schedule\Burns Switch Upgrade schedule.xlsx'

foreach ($row in $exdata) {

    if ($row.RoomID) {

        $FolderName = $rootDir + '\' + $row.RoomID.tolower()
        $FileName = $FolderName + '\cisco' + $row.DeviceName + '-confg-dump'
        if (Get-Item -Path $FolderName -ErrorAction Ignore) {

            Write-Host "Folder $($FolderName) Exists"
            #Create new file
            if (Get-Item -Path $FileName -ErrorAction Ignore) {
               # Write-Host -ForegroundColor green "File $($FileName) exists in $($FolderName)"
            }
            else {

                Write-Host -ForegroundColor red  "File $($FileName) missing from $($FolderName)"
            }
        }
    }
}

