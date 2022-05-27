# read the excell file and create the folders if they are missing from the encrypted volume.

#vars
$rootDir = 'y:\backups'

#import the excel file
$exData = Import-Excel 'C:\Users\ahall\Southern Alberta Institute of Technology\Campus Network Upgrade Project - General\OCM\Schedule\Burns Switch Upgrade schedule.xlsx'

foreach ($row in $exdata) {


    if ($row.RoomID) {
        #roomid is not empty
        $FolderName = $rootDir + '\' + $row.RoomID.tolower()
        if (Get-Item -Path $FolderName -ErrorAction Ignore) {

            Write-Host "$($FolderName) Folder Exists"

        }
        else {
            Write-Host "Folder Doesn't Exist, creating folder $($FolderName)"

            #PowerShell Create directory if not exists
            New-Item $FolderName -ItemType Directory
        }
    }

}