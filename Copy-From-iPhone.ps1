param(
[Parameter(Mandatory=$TRUE)]
[ValidateScript({Test-Path -IsValid $_})]
[string]$destPath
)

$sourceShellApplication = New-Object -com Shell.Application
$thisPCNameSpace = $sourceShellApplication.NameSpace(0x11)
$thisPCItems = $thisPCNameSpace.Items()

$iPhone = $thisPCItems | Where-Object Name -eq "Apple iPhone"
if ($iPhone -eq $null) {
    Write-Error "iPhone not found."
    return
}

$internalStorage = $iPhone.GetFolder.Items() | Where-Object Name -eq "Internal Storage"
if ($internalStorage -eq $null) {
    Write-Error "Internal storage not found"
    return
}

$sourceFolders = $internalStorage.GetFolder.Items()
$newFolderCount = $sourceFolders.Count
if ($newFolderCount -eq 0) {
    Write-Error "No folders found"
    return
}

if (-Not (Test-Path $destPath)) {
    New-Item $destPath -ItemType Directory
}
$destShellApplication = New-Object -com Shell.Application
$destNameSpace = $destShellApplication.NameSpace($destPath)

$statusFilePath = Join-Path $destPath "status.csv"
if (Test-Path $statusFilePath) {
    $statusObjects = Import-Csv -Path $statusFilePath
} else {
    $statusObjects = @()
}

$folderCounter = 0
$folderProgressParameters = @{
    Activity         = 'Folders'
    Id               = 0
    Status           = 'Folders Progress->'
}
foreach($sourceFolder in $sourceFolders) {
    $newFolderName = $sourceFolder.Name

    $statusObject = $statusObjects | Where-Object { $_.FolderName -eq $newFolderName }

    if ($statusObject.AllSuccess) {
        $folderCounter++
        $folderProgressParameters.CurrentOperation = "Skipped $newFolderName"
        $folderProgressParameters.PercentComplete = $folderCounter / $newFolderCount * 100
        Write-Progress @folderProgressParameters
        continue
    } else {
        $folderProgressParameters.CurrentOperation = "Copying $newFolderName"
        $folderProgressParameters.PercentComplete = $folderCounter / $newFolderCount * 100
        Write-Progress @folderProgressParameters
    }

    $destNameSpace.NewFolder($newFolderName)
    # creation method was void
    $newFolder = $destNameSpace.Items() | Where-Object Name -eq $newFolderName
    Write-Output $newFolder.Path

    $itemCount = $sourceFolder.GetFolder.Items().Count
    $itemCounter = 0
    $numSuccess = 0
    $numError = 0
    $numDuplicate = 0

    if ($itemCount -eq 0) {
        Write-Output "$newFolderName with 0 item count"
    }

    $items = $sourceFolder.GetFolder.Items()
    foreach ($item in $items) {
        $itemName = $item.Name

        $newItemPath = Join-Path $newFolder.Path $itemName
        if (-Not (Test-Path $newItemPath)) {
            try {
                # flag doesnt work for some reason
                $newFolder.GetFolder.CopyHere($item, 0x14)
                $numSuccess++
            } catch [Exception]{
                $numError++
                Write-Output "$newFolderName/$itemName error-ed out"
            }
        } else {
            $numDuplicate++
        }

        $itemCounter++
        if ($itemCount -eq 0) {
            $pctComplete = 99
        } else {
            $pctComplete = $itemCounter / $itemCount * 100
        }

        $itemProgressParameters = @{
            Activity         = 'Items'
            CurrentOperation = $itemName
            Id               = 1
            ParentId         = 0
            PercentComplete  = $pctComplete
            Status           = 'Items Progress->'
        }
        Write-Progress @itemProgressParameters
    }

    $statusRecord = [PSCustomObject]@{
        Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        FolderName = $sourceFolder.Name
        NumTotal = $itemCount
        NumSuccess = $numSuccess
        NumError = $numError
        NumDuplicate = $numDuplicate
        AnyDuplicate = $numDuplicate -gt 0
        AnyError = $numError -gt 0
        AllSuccess = $numSuccess -eq $itemCount
    }

    $folderCounter++
    $statusRecord | Export-Csv -Path $statusFilePath -Append -NoTypeInformation
}

$folderProgressParameters.CurrentOperation = "Done"
$folderProgressParameters.PercentComplete  = 100
Write-Progress @folderProgressParameters
