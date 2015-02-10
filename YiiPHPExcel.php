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
     * PHPExcel Settings here
     */
    public function __construct() {
        parent::__construct();
        //Use only one sheet
        $this->setActiveSheetIndex(0);
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
     * @return will return file to browser
     */
    public function createExcel($models, $attributes) {
        $letters = $this->getLetters($attributes);
        $this->setFirstRow($models, $attributes, $letters);
        $this->setRows($models, $attributes, $letters);
        $this->sendToBrowser();
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
     */
    private function setRows($models, $attributes, $letters) {
        $rownum = 2; //start from second row
        $counter = 0;
        foreach ($models as $model) {
            foreach ($attributes as $attribute) {
                $value = is_array($attribute) ? $this->getDataExcelValue($model, key($attribute)) :
                        $this->getDataExcelValue($model, $attribute);
                //Write value to row cell
                $this->getActiveSheet()->setCellValue($letters[$counter] . $rownum, $value);
                $counter++;
            }
            $counter = 0;
            $rownum++;
        }
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
     */
    private function setFirstRow($models, $attributes, $letters) {
        $counter = 0;
        foreach ($attributes as $attribute) {
            $label = is_array($attribute) ? current($attribute) : $models[0]->getAttributeLabel($attribute);
            $this->getActiveSheet()->setCellValue($letters[$counter] . '1', $label)
                    ->getStyle($letters[$counter] . '1')->getFont()->setBold(true);
            $counter++;
        }
    }

    /**
     * Will send the constructed excel to browser and end yii application
     */
    private function sendToBrowser() {
        // ** Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="gpd.xlsx"');
        header('Cache-Control: max-age=0');

        $objWriter = PHPExcel_IOFactory::createWriter($this, 'Excel2007');
        $objWriter->save('php://output');
        Yii::app()->end();
    }

}
