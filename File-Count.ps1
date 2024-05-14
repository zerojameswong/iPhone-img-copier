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

$items = $internalStorage.GetFolder.Items()
$folderCount = $items.Count
if ($folderCount -eq 0) {
    Write-Error "No folders found"
    return
}

Write-Output "Folder count is $folderCount"

$totalCount = 0
$currentFolderIdx = 0
foreach($sourceFolder in $items) {
    $itemCount = $sourceFolder.GetFolder.Items().Count
    $totalCount += $itemCount
    $name = $sourceFolder.Name
    $currentFolderIdx++
    Write-Output "#$currentFolderIdx : $name has $itemCount items, current total: $totalCount"
}
