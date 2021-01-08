<?php

# Cargar librerias y cosas necesarias
require_once "vendor/autoload.php";

# Indicar que usaremos el IOFactory
use PhpOffice\PhpSpreadsheet\IOFactory;

// Ruta y carga del archivo excel
$rutaArchivo = "banco.xlsx";
$documento = IOFactory::load($rutaArchivo);

$array = array();
$FilaArray = array();
$DateArray = array();
$RefArray = array();

# obtener conteo e iterar
$totalDeHojas = $documento->getSheetCount();
// echo "<h3>Total de hojas $totalDeHojas</h3>";

# Iterar hoja por hoja
for ($indiceHoja = 0; $indiceHoja < $totalDeHojas; $indiceHoja++) {

    # Obtener hoja en el índice que vaya del ciclo
    $hojaActual = $documento->getSheet($indiceHoja);
    // echo "<h3>Vamos en la hoja con índice $indiceHoja</h3>";

    # Iterar filas
    foreach ($hojaActual->getRowIterator() as $fila) {
        foreach ($fila->getCellIterator() as $celda) {
            # Formateado por ejemplo como dinero o con decimales
            $valorFormateado = $celda->getFormattedValue();

            # Fila, que comienza en 1, luego 2 y así...
            $fila = $celda->getRow();
            # Columna, que es la A, B, C y así...
            $columna = $celda->getColumn();

            // Expresion regular para que sea montos positivos.
            $exp = '/^[1-9][\.\d]*(,\d+)?$/';

            // Se verifica cuales son montos y cuales son positivos
            if ($columna == "D" && ($valorFormateado != null || $valorFormateado != "")) {
                if(preg_match($exp, $valorFormateado)) {
                    array_push($array, $valorFormateado);
                    array_push($FilaArray, $fila);
                }
            }
        }
    }
}

for ($indiceHoja = 0; $indiceHoja < $totalDeHojas; $indiceHoja++) {
    $hojaActual = $documento->getSheet($indiceHoja);
    $it = 0;
    $it2 = 0;
    foreach ($hojaActual->getRowIterator() as $fila) {
        foreach ($fila->getCellIterator() as $celda) {
            $valorFormateado = $celda->getFormattedValue();

            $fila = $celda->getRow();
            $columna = $celda->getColumn();

            // Expresiones regulares para fecha y para numeros positivos.
            $exp2 = '/^([0-2][0-9]|3[0-1])(\/|-)(0[1-9]|1[0-2])\2(\d{4})$/';
            $exp3 = '/^\d+$/';

            // Se verifica cuales son fechas y que sean fechas
            if(count($FilaArray)-1 >= $it){
                if ($columna == "A" && $fila == $FilaArray[$it] &&($valorFormateado != null || $valorFormateado != "")) {
                    if(preg_match($exp2, $valorFormateado)) {
                        array_push($DateArray, $valorFormateado);
                        $it++;
                    }
                }
            }
            
            // Se verifica cuales son referencias y que sean numeros
            if(count($FilaArray)-1 >= $it2){
                if ($columna == "B" && $fila == $FilaArray[$it2] &&($valorFormateado != null || $valorFormateado != "")) {
                    if(preg_match($exp3, $valorFormateado)) {
                        array_push($RefArray, $valorFormateado);
                        $it2++;
                    }
                }
            }
        }
    }
}
    // Arreglo de objetos para los movimientos
    $movArray = array();

    for ($i=0; $i < count($FilaArray); $i++) { 
        $movimiento = new stdClass();
        $movimiento->fecha = $DateArray[$i];
        $movimiento->ref = $RefArray[$i];
        $movimiento->monto = $array[$i];

        array_push($movArray, $movimiento);
    }

    for ($i=0; $i < count($movArray); $i++) { 
        echo $movArray[$i]->fecha, " ", $movArray[$i]->ref, " ", $movArray[$i]->monto, "<br><br>";
    }