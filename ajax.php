<?php

require 'vendor/autoload.php';
require 'Object.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();

$inputFileName = './stu_tb_trc.xlsx';

/** Load $inputFileName to a Spreadsheet Object  **/
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);


$getAnotherValues = function ($listIndex, $cellIndex, $listData) use ($spreadsheet) {

    //var_dump($get['form']);

    $sheet = $spreadsheet->getSheet($listIndex);

    $cellName = $sheet->getCell($cellIndex)->getValue();

    $letter = preg_replace('/\d+/', '', $cellIndex);

    $colName = $sheet->getCell($letter . 1)->getValue();

    $highestColumn = $sheet->getHighestColumn();

    $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); // e.g. 5

    // var_dump($colName);

    $rowIndex = preg_replace('/\D/', '', $cellIndex);


    for ($i = 2; $i <= $highestColumnIndex; $i++) {

        //$colName = $sheet->getCellByColumnAndRow($i, 1)->getValue();


        if ($sheet->getCellByColumnAndRow($i, $rowIndex)->getValue() != null) {

            // $listData[$colName][] = $sheet->getCellByColumnAndRow($i, $rowIndex)->getValue();

            $vertical = true;

        } else {
            $vertical = false;
        }

    }


    //var_dump($sheet->getCellByColumnAndRow(6, 3)->getValue()->getPlainText());

    $horizontal = !$vertical;

    if ($colName == 'Наименование') {

        // var_dump( $cellName = $sheet->getCell($cellIndex)->getValue());

        for ($j = 1; $j <= 10; $j++) {

            $rowIndex++;

            // var_dump($rowIndex);


            for ($i = 1; $i <= $highestColumnIndex; $i++) {

                $colName = $sheet->getCellByColumnAndRow($i, 1)->getValue();

                $cellValue = $sheet->getCellByColumnAndRow($i, $rowIndex)->getValue();


                if ($i == 4 && $cellValue == null) {
                    break 2;
                }

                $index = Object::getIndex($colName);

                if ($colName == 'Наименование') {
                    $listData[$index][$colName][0] = $sheet->getCell($cellIndex)->getValue();
                } else if ($cellValue != null) {

                    if (is_string($cellValue)) {
                        $listData[$index][$colName][] = $cellValue;
                    }
                    if (is_object($cellValue)) {
                        $listData[$index][$colName][] = $cellValue->getPlainText();
                    }
                } else {
                    $listData[$index][$colName][] = null;
                }


            }

        }


    } else if ($colName == 'Объект' || $colName == 'Разработчик СТУ' || $colName == 'Заключение МЧС (номер, дата, исполнитель)') {

        for ($j = $rowIndex; $j < $rowIndex + 10; $j++) {

            for ($i = 1; $i <= $highestColumnIndex; $i++) {

                $colName = $sheet->getCellByColumnAndRow($i, 1)->getValue();

                $cellValue = $sheet->getCellByColumnAndRow($i, $j)->getValue();

                $index = Object::getIndex($colName);

                if ($colName == 'Наименование' && $j == $rowIndex && $rowIndex != 1) {
                    $listData[$index][$colName][$j] = $sheet->getCellByColumnAndRow($i, $rowIndex - 1)->getValue();


                } else {

                    if (is_string($cellValue)) {
                        $listData[$index][$colName][$j] = $cellValue;
                    }
                    if (is_object($cellValue)) {
                        $listData[$index][$colName][$j] = $cellValue->getPlainText();
                    }

                }


                if ($j != $rowIndex && ($colName == 'Объект' || $colName == 'Разработчик СТУ' || $colName == 'Заключение МЧС (номер, дата, исполнитель)') && $cellValue != NULL) {

                    foreach ($listData as &$v) {
                        foreach ($v as &$v2) {
                            unset($v2[$j - 1]);
                            unset($v2[$j]);
                        }
                    }

                    break 2;
                }
            }

        }

    } else {
        for ($i = 1; $i <= $highestColumnIndex; $i++) {
            $colName = $sheet->getCellByColumnAndRow($i, 1)->getValue();

            $cellValue = $sheet->getCellByColumnAndRow($i, $rowIndex)->getValue();

            $index = Object::getIndex($colName);

            if ($colName == 'Наименование') {
                for($j = $rowIndex; $j > 1; $j--) {
                    //var_dump($j);
                    if($sheet->getCellByColumnAndRow($i, $j)->getValue() != null) {
                        $listData[$index][$colName][$rowIndex] = $sheet->getCellByColumnAndRow($i, $j)->getValue();
                        break;
                    }
                }

            }else if($colName == 'Объект' || $colName == 'Разработчик СТУ' || $colName == 'Заключение МЧС (номер, дата, исполнитель)'){
                for($j = $rowIndex; $j > 1; $j--) {
                   // var_dump($j);
                     if($sheet->getCellByColumnAndRow($i, $j)->getValue() != null) {
                         $cellValue = $sheet->getCellByColumnAndRow($i, $j)->getValue();

                         if (is_string($cellValue)) {
                             $listData[$index][$colName][$rowIndex] = $cellValue;
                         }
                         if (is_object($cellValue)) {
                             $listData[$index][$colName][$rowIndex] = $cellValue->getPlainText();
                         }

                     }
                }
            } else {

                if (is_string($cellValue)) {
                    $listData[$index][$colName][$rowIndex] = $cellValue;
                }
                if (is_object($cellValue)) {
                    $listData[$index][$colName][$rowIndex] = $cellValue->getPlainText();
                }

            }

        }

    }


    // var_dump($cellIndex);


    return $listData;
};


if ($_GET['ajax']) {
    $arr = ['n' => 1];

    unset($_GET['ajax']);

    $listData = [];

    //var_dump($_GET['form']);

    foreach ($_GET['form'] as $val) {

        $listIndex = explode('_', $val['value'])[0];

        $cellIndex = explode('_', $val['value'])[1];


        $sheet = $spreadsheet->getSheet($listIndex);

        $cellName = $sheet->getCell($cellIndex)->getValue();

        $listName = $val['name'];

        //$listData[0][$listName][] = $cellName;

    }

    $listData = $getAnotherValues($listIndex, $cellIndex, $listData);


//var_dump($listData);

    header('Content-Type: application/json');
    echo json_encode($listData);
}


?>
