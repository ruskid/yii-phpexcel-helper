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

Setting the start row of the data. and possibility to separate data in excel by blank line.
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

$excel->writeBlankRow(); // will write blank row which will visually separate OFERTAS AND DUDAS data.

$excel->writeRecordsToExcel($solicitud->DUDAS, [
    ['PROVEEDOR->NOMBRE' => Yii::t('app', 'Nombre de proveedor')],
    'PREGUNTA',
    'RESPUESTA',
    ['getEstadoLabel' => Yii::t('app', 'Estado')],
    ['getAmbitoLabel' => Yii::t('app', 'Ambito')],
]);
$excel->sendToBrowser();
```


You can also set the style for the data.
--------------------------
```php

$firstRowStyle = array(
    'font' => array(
        'bold' => true,
    ),
);

$allRowsStyle = array(
    'font' => array(
        'name' => 'Arial',
        'color' => array(
            'rgb' => '333333'
        )
    ),
);
$excel->writeRecordsToExcel($compras, [
    ['TITULO' => Yii::t('app', 'Título Compra')],
    ['JUSTIFICACION' => Yii::t('app', 'Justificación')],
    ['DESCRIPCION' => Yii::t('app', 'Descripción')],
    ['getEstadoLabel' => Yii::t('app', 'Estado')],
    ['FECHA' => Yii::t('app', 'Fecha')],
    ['getLoadSolicitudLabel' => Yii::t('app', 'Loa o Solicitud')],
    ['USER->NOMBRE' => Yii::t('app', 'Usuario')],
], $firstRowStyle, $allRowsStyle);

$excel->sendToBrowser();
```