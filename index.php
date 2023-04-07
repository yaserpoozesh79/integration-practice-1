<?php

require_once __DIR__."\\vendor\\autoload.php";

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

$phpWord = IOFactory::load('description.docx');

/* Note: any element you append to a document must reside inside of a Section. */

// Adding an empty Section to the document...
$section = $phpWord->addSection();

// Adding Text element with font customized using named font style...
$fontStyleName = 'oneUserDefinedStyle';
$phpWord->addFontStyle(
    $fontStyleName,
    array('name' => 'shabnam FD', 'size' => 30, 'color' => '1B2232', 'bold' => true)
);
$section->addText(
    "سلام. این یک متن آزمایشی است از نحوه کارکرد کتابخانه phpword.",
    $fontStyleName
);


$i=2;
while(file_exists($fileName = $i.'.docx')){
    $i++;
}


// Saving the document as OOXML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save($fileName);

echo 'done. file created: '.$fileName;
