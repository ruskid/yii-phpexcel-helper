# Yii1 PHPExcel Helper.
Can extract CActiveRecords to excel attachment file.

If label wasn't specified it will take getAttributeLabel of the Active Record.

Usage
--------------------------
```php
$compras = COMPRA::model()->findAll();
$excel = new YiiPHPExcel;
return $excel->createExcel($compras, [
    'TITLECOMPRA',
    ['USER->NOMBRE' => Yii::t('app', 'Usuario')],
    'FECHAINICIO',
    'FECHAFIN',
    'UNIDAD',
    'getProductoLabel',
    'PROVEEDOR->NOMBRE',
    ['getEstadoLabel' => Yii::t('app', 'Estado')]
]);
```
