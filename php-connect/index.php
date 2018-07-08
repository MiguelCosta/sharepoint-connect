<?php

require_once './vendor/autoload.php';
use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\SharePoint\ClientContext;
use Office365\PHP\Client\SharePoint\File;
use Office365\PHP\Client\SharePoint\SPList;
require_once './vendor/vgrem/php-spo/tests/ListExtensions.php';

try {
    $Url = "https://mpcjellycode.sharepoint.com/sites/Site";
    $UserName = "mcosta@mpcjellycode.onmicrosoft.com";
    $Password = "JellyCode2018";

    $authCtx = new AuthenticationContext($Url);
    $authCtx->acquireTokenForUser($UserName, $Password); //authenticate

    $ctx = new ClientContext($Url, $authCtx);

    $localPath = "./downloadFolder/"; // Folder should be exist in ..\sharepoint-connect\php-connect\downloadFolder
    $targetLibraryTitle = "Documents";
    $targetFolderUrl = "/sites/Site/Shared Documents";

    echo "<h2>FOLDERS</h2>";
    $list = ListExtensions::ensureList($ctx->getWeb(), $targetLibraryTitle, \Office365\PHP\Client\SharePoint\ListTemplateType::DocumentLibrary);
    enumFolders($list);

    echo "<h2>CREATE FOLDER</h2>";
    createSubFolder($ctx, $targetFolderUrl, "Folder2001");

    echo "<h2>GET FILE 1</h2>";
    $fileUrl = "/sites/Site/Shared Documents/catalogos_teste.xls";
    $file = $ctx->getWeb()->getFileByServerRelativeUrl($fileUrl);
    $ctx->load($file);
    $ctx->executeQuery();
    printFileProperties($file);

    echo "<h2>GET FILE 2</h2>";
    $folderUrl = "Shared Documents";
    $fileUrl = "Document.docx";
    $file = $ctx->getWeb()->getFolders()->getByUrl($folderUrl)->getFiles()->getByUrl($fileUrl);
    $ctx->load($file);
    $ctx->executeQuery();
    printFileProperties($file);

    echo "<h2>DOWNLOAD FILE</h2>";
    $fileUrl = "/sites/Site/Shared Documents/catalogos_teste.xls";
    downloadFile($ctx, $fileUrl, $localPath, "catalogos_teste.xls");

} catch (Exception $ex) {
    echo "<h1>ERRO</h1>";
    echo $ex->getMessage();
}

echo "<h2>END SCRIPT PHP</h2>";

function printFileProperties(File $file)
{
    echo "<ul>";
    echo "<li>Name: '{$file->getProperty("Name")}'</li>";
    echo "<li>ServerRelativeUrl: '{$file->getProperty("ServerRelativeUrl")}'</li>";
    echo "<li>TimeCreated: '{$file->getProperty("TimeCreated")}'</li>";
    echo "<li>TimeLastModified: '{$file->getProperty("TimeLastModified")}'</li>";
    echo "<li>UIVersion: '{$file->getProperty("UIVersion")}'</li>";
    echo "<li>UIVersionLabel: '{$file->getProperty("UIVersionLabel")}'</li>";
    echo "<li>UniqueId: '{$file->getProperty("UniqueId")}'</li>";
    echo "</ul>";
    //print_r($file);
}

function enumFolders(SPList $list)
{
    $ctx = $list->getContext();
    $folders = $list->getRootFolder()->getFolders();

    if ($folders->getServerObjectIsNull() == true) { //determine whether folders has been loaded or not
        $ctx->load($folders);
        $ctx->executeQuery();
    }

    echo "<ol>";
    foreach ($folders->getData() as $folder) {
        print "<li>'{$folder->Name}'</li>";
    }
    echo "</ol>";
}

function createSubFolder(ClientContext $ctx, $parentFolderUrl, $folderName)
{
    $files = $ctx->getWeb()->getFolderByServerRelativeUrl($parentFolderUrl)->getFiles();
    $ctx->load($files);
    $ctx->executeQuery();
    //print files info
    /* @var $file \Office365\PHP\Client\SharePoint\File */
    echo "<h3>Current Items</h3>";
    echo "<ol>";
    foreach ($files->getData() as $file) {
        print "<li>File name: '{$file->getProperty("ServerRelativeUrl")}'</li>";
    }
    echo "</ol>";
    $parentFolder = $ctx->getWeb()->getFolderByServerRelativeUrl($parentFolderUrl);
    $childFolder = $parentFolder->getFolders()->add($folderName);
    $ctx->executeQuery();
    print "Child folder {$childFolder->getProperty("ServerRelativeUrl")} has been created ";
}

function downloadFile(ClientContext $ctx, string $fileUrl, string $targetFolderPath, string $fileName)
{
    $fileContent = File::openBinary($ctx, $fileUrl);
    $fileLocation = "{$targetFolderPath}{$fileName}";
    file_put_contents($fileLocation, $fileContent);
    print "File {$fileUrl} has been downloaded successfully to {$fileLocation}.";
}
