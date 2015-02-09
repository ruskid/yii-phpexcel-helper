# Yii1 PHPExcel Helper.
Can extract CActiveRecords to excel attachment file.

YiiPHPExcel is extended by PHPExcel, so you have to download PHPExcel to your extensions or vendor folder and import it in the main.php like so.
```php
'import' => array(
...
'application.components.YiiPHPExcel',
'ext.PHPExcel.PHPExcel',
...
),
```

If label wasn't specified it will take getAttributeLabel of the Active Record as an excel column header.

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
