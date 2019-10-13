<?php

echo $_POST['selects'][1]['name'];

/*

$phpWord = new  \PhpOffice\PhpWord\PhpWord();

$phpWord->setDefaultFontName('Times New Roman');

$phpWord->setDefaultFontSize(14);

$properties = $phpWord->getDocInfo();

$properties->setCreator('Name');
$properties->setCompany('Company');
$properties->setTitle('Title');
$properties->setDescription('Description');
$properties->setCategory('My category');
$properties->setLastModifiedBy('My name');
$properties->setCreated(mktime(0, 0, 0, 3, 12, 2015));
$properties->setModified(mktime(0, 0, 0, 3, 14, 2015));
$properties->setSubject('My subject');
$properties->setKeywords('my, key, word');


$sectionStyle = array(

    'orientation' => 'landscape',
    'marginTop' => \PhpOffice\PhpWord\Shared\Converter::pixelToTwip(10),
    'marginLeft' => 600,
    'marginRight' => 600,
    'colsNum' => 1,
    'pageNumberingStart' => 1,
    'borderBottomSize'=>100,
    'borderBottomColor'=>'C0C0C0'

);
$section = $phpWord->addSection($sectionStyle);

$sectionStyle = array(

    'orientation' => 'landscape',
    'marginTop' => \PhpOffice\PhpWord\Shared\Converter::pixelToTwip(10),
    'marginLeft' => 600,
    'marginRight' => 600,
    'colsNum' => 1,
    'pageNumberingStart' => 1,
    'borderBottomSize'=>100,
    'borderBottomColor'=>'C0C0C0'

);
$section = $phpWord->addSection($sectionStyle);

$text = "PHPWord is a library written in pure PHP that provides a set of classes to write to and read from different document file formats.";
$fontStyle = array('name'=>'Arial', 'size'=>36, 'color'=>'075776', 'bold'=>TRUE, 'italic'=>TRUE);
$parStyle = array('align'=>'right','spaceBefore'=>10);

$section->addText(htmlspecialchars($text), $fontStyle,$parStyle);


$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord,'Word2007');
$objWriter->save('doc.docx');

*/