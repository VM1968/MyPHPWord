<?php
require_once 'Shared/Zip.php';
require_once 'Shared/Xml.php';

$wordFile='file2.docx';

$newFile='newFile.docx';
$tempDir='TempZipDir';
$zip=new Zip();
$result=$zip->open($wordFile);


if ($result) {
     $zip->extractTo($tempDir);
     $zip->close();
     //$zip->create($newFile);
     $xml=new Xml();
     $docx=$xml->document($tempDir);
     // $docx=$xml->relations($tempDir);
     // $docx=$xml->parsedocument($tempDir);
}


echo $docx;