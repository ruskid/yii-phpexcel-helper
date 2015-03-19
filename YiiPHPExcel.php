<?php

/**
 * @copyright Copyright Victor Demin, 2015
 * @license https://github.com/ruskid/yii-phpexcel-helper/LICENSE
 * @link https://github.com/ruskid/yii-phpexcel-helper#readme
 */

/**
 * Little PHPExcel helper for yii1
 * @author Victor Demin <demin@trabeja.com>
 * @author Jordi Freixa Serra
 */
class YiiPHPExcel extends PHPExcel {

    /**
     * The row number in excel from where to start writing the data.
     * Default: 1, first line.
     * @var int
     */
    private $_startRowNumber = 1;

    /**
     * PHPExcel Settings here
     */
    public function __construct() {
        parent::__construct();
        //Use only one sheet
        $this->setActiveSheetIndex(0);
    }

    /**
     * Will set the start row number
     * @param integer $startFromRow
     */
    public function setStartRowNumber($startRowNumber) {
        $this->_startRowNumber = $startRowNumber;
    }

    /**
     * Will get the start row number
     * @return integer
     */
    public function getStartRowNumber() {
        return $this->_startRowNumber;
    }

    /**
     * Will increment the row number, so the next data could be written after the blank row.
     */
    public function writeBlankRow() {
        $rownum = $this->getStartRowNumber();
        $this->setStartRowNumber(++$rownum);
    }

    /**
     * Will construct PHPExcel. You can specify attributes, methods and relations in array of attributes.
     *
     * Example of usage:
     * $excel = new YiiPHPExcel;
     * return $excel->createExcel($compras, [
     *               'TITLECOMPRA',
     *               ['USER->NOMBRE' => Yii::t('app', 'Usuario')],
     *               'FECHAINICIO',
     *               'FECHAFIN',
     *               'UNIDAD',
     *               'getProductoLabel',
     *               'PROVEEDOR->NOMBRE',
     *               ['getEstadoLabel' => Yii::t('app', 'Estado')]
     * ]);
     *
     * @param array $models Array of CActiveRecords
     * @param array $attributes Array of attributes, methods and relations.
     * @param array $firstRowStyle Array of styles for the first row (header)
     * @param array $allRowsStyle Array of styles for data rows
     * @return will return file to browser
     */
    public function writeRecordsToExcel($models, $attributes, $firstRowStyle = ['font' => ['bold' => true]], $allRowsStyle = []) {
        if ($models && $attributes) {
            $letters = $this->getLetters($attributes);
            $this->setFirstRow($models, $attributes, $letters, $firstRowStyle);
            $this->setRows($models, $attributes, $letters, $allRowsStyle);
        }
    }

    /**
     * Will get exact number of letter needed for the row
     * @param array $array Will get array of first column row
     * @return array Will get array of letters
     */
    private function getLetters($array) {
        $letters = [];
        $count = 0;
        $maxCount = count($array);
        for ($char = 'A'; $char <= 'Z' && $count < $maxCount; $char++) {
            array_push($letters, $char);
            $count++;
        }
        return $letters;
    }

    /**
     * Will set values to the excel
     * @param CActiveRecord[] $models
     * @param array $attributes
     * @param array $letters
     * @param array $style
     */
    private function setRows($models, $attributes, $letters, $style) {
        $rownum = $this->getStartRowNumber();
        $counter = 0;
        if (!is_array($models)) {
            $models = [$models];
        }
        foreach ($models as $model) {
            foreach ($attributes as $attribute) {
                $value = is_array($attribute) ? $this->getDataExcelValue($model, key($attribute)) :
                        $this->getDataExcelValue($model, $attribute);
                //Write value to row cell
                $this->getActiveSheet()->setCellValue($letters[$counter] . $rownum, $value)
                        ->getStyle($letters[$counter] . $rownum)->applyFromArray($style);
                $counter++;
            }
            $counter = 0;
            $rownum++;
        }
        //After writing the data into excel, set the next start row number.
        $this->setStartRowNumber($rownum);
    }

    /**
     * Will get value for model by attribute type. (property, method, relation)
     * @param CActiveRecord $model
     * @param string $attribute
     * @return string Value
     */
    private function getDataExcelValue($model, $attribute) {
        if ($model->hasAttribute($attribute)) {
            return $model->getAttribute($attribute);
        }
        if (method_exists($model, $attribute)) {
            return call_user_func(array($model, $attribute));
        }
        //If relation. Can be a chain of relations  //USER -> PROVEEDOR -> NOMBRE
        if (strpos($attribute, '->') !== false) {
            $arr = explode('->', $attribute);
            foreach ($arr as $attribute) {
                if ($model != null) {
                    $relations = $model->relations();
                    if (isset($relations[$attribute])) {//If is a chain relation. then recursive.
                        $model = $model->getRelated($attribute);
                    } elseif ($model->hasAttribute($attribute)) {//if is attribute
                        return $model->getAttribute($attribute);
                    } elseif (method_exists($model, $attribute)) {//if is method
                        return call_user_func(array($model, $attribute));
                    }
                }
            }
        }
    }

    /**
     * Will set first row in excel with labels of CActiveRecord or labels in attributes array
     * @param CActiveRecord[] $models
     * @param array $attributes
     * @param array $letters
     * @param array $style
     */
    private function setFirstRow($models, $attributes, $letters, $style) {
        $rownum = $this->getStartRowNumber();
        $counter = 0;
        foreach ($attributes as $attribute) {
            $label = is_array($attribute) ? current($attribute) : $models[0]->getAttributeLabel($attribute);
            $this->getActiveSheet()->setCellValue($letters[$counter] . $rownum, $label)
                    ->getStyle($letters[$counter] . $rownum)->applyFromArray($style);
            $counter++;
        }
        //After writing the data into excel, set the next start row number.
        $this->setStartRowNumber( ++$rownum);
    }

    /**
     * Will send the constructed excel to browser and end yii application
     * @param string $filename
     */
    public function sendToBrowser($filename = 'file') {
        // ** Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
        header("Pragma: "); //IE8 quick fix.
        header("Cache-Control: "); //IE8 quick fix.

        $objWriter = PHPExcel_IOFactory::createWriter($this, 'Excel2007');
        $objWriter->save('php://output');
        Yii::app()->end();
    }

}
