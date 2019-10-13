<?php

class Object
{
    public static function getIndex($colName)
    {
        switch ($colName) {
            case 'Наименование':
                $index = 1;
                break;
            case 'Объект':
                $index = 2;
                break;
            case 'Разработчик СТУ':
                $index = 3;
                break;
            case 'Заключение МЧС (номер, дата, исполнитель)':
                $index = 4;
                break;
            case 'Необходимость разработки СТУ':
                $index = 5;
                break;
            case 'Примечание':
                $index = 6;
                break;
            case 'пункт в СТУ':
                $index = 7;
                break;
            case 'Норматив':
                $index = 8;
                break;

        }

        return $index;
    }
}