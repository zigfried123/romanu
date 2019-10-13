<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$spreadsheet = new Spreadsheet();

$inputFileName = './files/file.xlsx';


if(!empty($_FILES)){

    $uploaddir = './files/';
    $uploadfile = $uploaddir . 'file.xlsx';

    //move_uploaded_file($_FILES['f']['tmp_name'], $uploadfile);


   // echo '<pre>';
    if (move_uploaded_file($_FILES['f']['tmp_name'], $uploadfile)) {
       // echo "Файл корректен и был успешно загружен.\n";
    } else {
       // echo "Возможная атака с помощью файловой загрузки!\n";
    }


    header('Location: /');

}


try {

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);


$sheetCount = $spreadsheet->getSheetCount();

//$worksheet->getHighestRow(); // e.g. 10
//$highestColumn = $worksheet->getHighestColumn();


$sheet = $spreadsheet->getSheet(1);


$getListData = function ($letter) use ($spreadsheet, $sheetCount) {

    $data = [];

    for ($i = 0; $i < $sheetCount; $i++) {
        $sheet = $spreadsheet->getSheet($i);

        $sheetData = $sheet->toArray(null, true, true, true);

        $highestRow = $sheet->getHighestRow();

        $highestColumnIndex = $sheet->getHighestColumn();


        for ($row = 2; $row <= $highestRow; ++$row) {

            if ($sheet->getCell("$letter$row")->getValue() == null) continue;

            //var_dump($sheet->getCodeName());

            $listIndex = explode('_', $sheet->getCodeName())[1];

            $listIndex = intval($listIndex);

            //var_dump($listIndex);

            $data[$listIndex . '_' . $letter . $row] = ['value' => $sheet->getCell($letter . $row)->getValue()];


        }


    }

    return $data;

};


//var_dump($sheet->getCellByColumnAndRow(2, 3)->getValue());

//var_dump($getListData('A')); die;

$renderList = function ($letter, $selectName) use ($sheet, $getListData) {

    $listData = $getListData($letter);

    echo $sheet->getCell($letter . 1)->getValue() . '<br>';


    ?>


    <select name="<?= $selectName ?>">
        <option disabled selected value>Не выбрано</option>
        <?php foreach ($listData as $key => $val) {

            $selected = $_GET[$selectName] == $key ? 'selected' : '';

            ?>

            <option <?= $selected ?> value='<?= $key ?>'>
                <?= mb_substr($val['value'], 0, 900); ?>
            </option>
        <?php } ?>
    </select><br>


    <?php

};

?>

<form>

    <?php

    $renderList('A', 'наименование');

    $renderList('F', 'объект');

    $renderList('G', 'разработчик');

    $renderList('H', 'заключение');

    $renderList('B', 'необходмость');

    $renderList('C', 'примечание');

    $renderList('D', 'пункт сгу');

    $renderList('E', 'норматив');



    ?>


    </form>

    <br>
    <button type="submit">Отправить</button><br><br>

    <button onclick="window.location.reload();">сброс</button>

    <?php
}catch(Exception $e) {

    ?>

    <form enctype="multipart/form-data" method="post">
        <p><input type="file" name="f">
            <input type="submit" value="Отправить"></p>
    </form>

    <?php

}


/*
       echo '<tr>' . PHP_EOL;
       for ($col = 1; $col <= $highestColumnIndex; ++$col) {
           $value = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
           echo '<td>' . $value . '</td>' . PHP_EOL;
       }
       echo '</tr>' . PHP_EOL;
       */


//$cellValue = $spreadsheet->getActiveSheet()->getCell('A1')->getValue();


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

?>

<div id="test"></div>

<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>

<script>


    $('select').change(function () {

        $.ajax({
            url: 'ajax.php',
            data: {form: $('form').serializeArray(), ajax: true},
            dataType: 'json',
            success: function (data) {

                //console.log(data);

                $('form').empty();

                //$('')

                let i = 0;

                $.each(data, function (k, v) {
                    i++;

                    let key = Object.keys(v)[0];

                   // console.log(Object.keys(v)[0]);

                    $('form').append('<span>'+Object.keys(v)[0]+'</span><br>');

                    $('form').append('<select id="'+i+'"></select><br>');

                    console.log(v);



                    $.each(v[key], function (k2, v2) {

                        if(v2 != null && v2 != '') {

                            $('select#'+i).append('<option value="' + v2 + '">' + v2 + '</option>');

                        }else{
                          //  $('select#'+i).remove();
                        }

                        //console.log(v2);

                    });
                });

                //$('select').option('');

               // console.log(data.form);
                /*
                                $.each(data.form,function(k,v){
                                    console.log(v);
                                });
                                */

            }
        });


    });

</script>

