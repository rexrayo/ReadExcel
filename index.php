<?php
require 'Excel.php';?>

<h1>El Excel</h1>

<?php

$rutaArchivo = "banco.xlsx";
$excel = new Excel($rutaArchivo); 
echo $excel->drotacaExcel();?>