<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$spreadsheet = new Spreadsheet();

$inputFileName = './files/file.xlsx';


if(!empty($_FILES['f'])){

    $uploaddir = './files/';
    $uploadfile = $uploaddir . 'file.xlsx';

    //move_uploaded_file($_FILES['f']['tmp_name'], $uploadfile);


   // echo '<pre>';
    if (move_uploaded_file($_FILES['f']['tmp_name'], $uploadfile)) {
       // echo "Файл корректен и был успешно загружен.\n";
    } else {
       // echo "Возможная атака с помощью файловой загрузки!\n";
    }


    //header('Location: /');

}




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
                <?= mb_substr($val['value'], 0, 300); ?>
            </option>
        <?php } ?>
    </select><br>


    <?php

};

?>

<form method="post" id="main-form">

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


    <button onclick="window.location.reload();">сброс</button>

    <?php


?>

<form enctype="multipart/form-data" method="post">
    <p><input type="file" name="f">
        <input type="submit" value="Загрузить"></p>
</form>

<?php


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



*/

?>

<div id="test"></div>

<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>

<script>


    $('select').change(function () {

       // console.log($('form#main-form'));

     //   console.log($('form#main-form').serializeArray());

        $.ajax({
            url: 'ajax.php',
            data: {form: $('form#main-form').serializeArray(), ajax: true},
            dataType: 'json',
            success: function (data) {

                //console.log(data);

                $('form#main-form').empty();

                //$('')

                let i = 0;

                $.each(data, function (k, v) {
                    i++;

                    let key = Object.keys(v)[0];

                   // console.log(Object.keys(v)[0]);

                    $('form#main-form').append('<span>'+Object.keys(v)[0]+'</span><br>');

                    $('form#main-form').append('<input type="checkbox"><select style="width:1200px" id="'+i+'"></select><br>');

                   // console.log(v);



                    $.each(v[key], function (k2, v2) {

                        if(v2 != null && v2 != '') {

                            $('select#'+i).append('<option value="' + v2 + '">' + v2 + '</option>');

                        }else{
                          //  $('select#'+i).remove();
                        }

                        //console.log(v2);

                    });
                });

                $('form#main-form').append('<br><button type="submit">Отправить</button>');

                $('button[type="submit"]').click(function (e){
                    e.preventDefault();

                    let selects = [];

                    $('input[type="checkbox"]').each(function(k,v){


                        if($(v).is(':checked')){

                            selects.push({name:  $(v).prevAll('span').html(), val: $(v).next().val()});

                        }


                    });

                    $.post('word.php',{'selects': selects})

                        .done(function(data){
                           // console.log(data);
                        })

                })



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

