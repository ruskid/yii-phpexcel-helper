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

If label wasn't specified it will take getAttributeLabel of the Active Record as an excel column header. You can rewrite sendToBrowser method for excel extension and stuff.

Usage
--------------------------
```php
$solicitudes = SOLICITUD::model()->findAll($criteria);
$excel = new YiiPHPExcel;
$excel->writeRecordsToExcel($solicitudes, [
            ['ID_REFERENCIA' => Yii::t('app', 'Referencia')],
            'NOMBRE_SOLICITUD',
            ['PERFIL_PUESTO->JOBROLE' => Yii::t('app', 'Jobrole')],
            'DURACION',
            'F_INCORPORACION',
            ['getEstadoLabel' => Yii::t('app', 'Estado')],
            ['getCountTotalOfertas' => Yii::t('app', 'Total de ofertas')],
            ['getCountNuevasOfertas' => Yii::t('app', 'Ofertas por ver')],
]);
$excel->sendToBrowser();
```

Setting the start row of the data.
--------------------------
```php
$excel = new YiiPHPExcel;
$excel->getActiveSheet()->setCellValue('A1', 'Other info');
$excel->getActiveSheet()->setCellValue('A2', 'Other info2');

$excel->setStartRowNumber(4); //start to write into excel from the row #4
$excel->writeRecordsToExcel($solicitud->OFERTAS, [
    ['ID_REFERENCIA' => Yii::t('app', 'Referencia oferta')],
    ['PROVEEDOR->NOMBRE' => Yii::t('app', 'Nombre de proveedor')],
    ['REF_CANDIDATO' => Yii::t('app', 'Referencia candidato')],
    'PRECIO_HORA',
    'INFO_ADICIONAL',
    ['getEstadoLabel' => Yii::t('app', 'Estado')],
]);
$excel->writeRecordsToExcel($solicitud->DUDAS, [
    ['PROVEEDOR->NOMBRE' => Yii::t('app', 'Nombre de proveedor')],
    'PREGUNTA',
    'RESPUESTA',
    ['getEstadoLabel' => Yii::t('app', 'Estado')],
    ['getAmbitoLabel' => Yii::t('app', 'Ambito')],
]);
$excel->sendToBrowser();
```

