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

$newFolderCount = $internalStorage.GetFolder.Items().Count
if ($newFolderCount -eq 0) {
    Write-Error "No folders found"
    return
}


if (-Not (Test-Path $destPath)) {
    New-Item $destPath -ItemType Directory
}
$destShellApplication = New-Object -com Shell.Application
$destNameSpace = $destShellApplication.NameSpace($destPath)

$completedFilePath = Join-Path $destPath ".completed"
$completed = @{}
if (Test-Path $completedFilePath) {
    Get-Content $completedFilePath | ForEach-Object {
        $completed[$_] = $true
    }
}

$folderCounter = 0
$folderProgressParameters = @{
    Activity         = 'Folders'
    Id               = 0
    Status           = 'Folders Progress->'
}
foreach($sourceFolder in $internalStorage.GetFolder.Items()) {
    $newFolderName = $sourceFolder.Name

    if ($completed[$newFolderName]) {
        $folderCounter++
        $folderProgressParameters.CurrentOperation = "Skipped $newFolderName"
        $folderProgressParameters.PercentComplete  = $folderCounter / $newFolderCount * 100
        Write-Progress @folderProgressParameters
        continue
    } else {
        $folderProgressParameters.CurrentOperation = "Copying $newFolderName"
        $folderProgressParameters.PercentComplete  = $folderCounter / $newFolderCount * 100
        Write-Progress @folderProgressParameters
    }

    $destNameSpace.NewFolder($newFolderName)
    # creation method was void
    $newFolder = $destNameSpace.Items() | Where-Object Name -eq $newFolderName

    $itemCount = $sourceFolder.GetFolder.Items().Count
    $itemCounter = 0
    foreach ($item in $sourceFolder.GetFolder.Items()) {
        $itemName = $item.Name

        $newItemPath = Join-Path $newFolder.Path $itemName
        if (-Not (Test-Path $newItemPath)) {
            # flag doesnt work for some reason
            $newFolder.GetFolder.CopyHere($item, 0x14)
        }

        $itemCounter++
        $itemProgressParameters = @{
            Activity         = 'Items'
            CurrentOperation = $itemName
            Id               = 1
            ParentId         = 0
            PercentComplete  = $itemCounter / $itemCount * 100
            Status           = 'Items Progress->'
        }
        Write-Progress @itemProgressParameters
    }

    $folderCounter++
    Add-Content -path $completedFilePath -value $newFolderName
}

$folderProgressParameters.CurrentOperation = "Done"
$folderProgressParameters.PercentComplete  = 100
Write-Progress @folderProgressParameters
