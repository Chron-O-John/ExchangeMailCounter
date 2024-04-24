connect-exchangeOnline
$emailaddress = Read-Host -Prompt 'E-Mail-adresse des Kontos'
$folderscope = Read-Host -Prompt 'Welcher Folder? (oberste Ebene)...leer für Inbox'
if ([string]::IsNullOrWhiteSpace($folderscope)){
	$folderscope = 'Inbox'
}

$itemsnum = Get-MailboxFolderStatistics -Identity $emailaddress -FolderScope $folderscope -ResultSize unlimited| Select ItemsInFolder, VisibleItemsInFolder, HiddenItemsInFolder, DeletedItemsInFolder, ItemsInFolderAndSubfolders

$itemsInFolder = 0
$visibleItemsInFolder = 0
$hiddenItemsInFolder = 0
$deletedItemsInFolder = 0

foreach ($element in $itemsnum) {
	$itemsInFolder += $element.ItemsInFolder
	$visibleItemsInFolder += $element.VisibleItemsInFolder
	$hiddenItemsInFolder += $element.HiddenItemsInFolder
	$deletedItemsInFolder += $deletedItemsInFolder
}

echo("1st Element ItemsInFolderAndSubfolders:`t`t" + $itemsnum[0].ItemsInFolderAndSubfolders)
echo("Items in Folders:`t" + $itemsInFolder)
echo("Visible Items in Folders:`t" + $visibleItemsInFolder)
echo("Hidden Items in Folders:`t" + $hiddenItemsInFolder)
echo("Deleted Items in Folders:`t" + $deletedItemsInFolder)

