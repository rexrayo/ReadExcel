<?php

# Cargar librerias y cosas necesarias
require_once "vendor/autoload.php";

# Indicar que usaremos el IOFactory
use PhpOffice\PhpSpreadsheet\IOFactory;

class Excel 
{
    
    var $path, $documento; 
    var $dateC, $refC, $amountC;
    var $amountsArray = array();
    var $filasArray = array();
    var $DateArray = array();
    var $RefArray = array();
    var $totalDeHojas;
    var $movimientos = array();

    public function __construct($path)
    {
        $this->path = $path;
        $this->documento = IOFactory::load($this->path);
        $this->totalDeHojas = $this->documento->getSheetCount();

    }

    function columns(){
        for ($indiceHoja = 0; $indiceHoja < $this->totalDeHojas; $indiceHoja++) {

            $hojaActual = $this->documento->getSheet($indiceHoja);

            foreach ($hojaActual->getRowIterator() as $fila) {
                foreach ($fila->getCellIterator() as $celda) {
                    $valorFormateado = $celda->getFormattedValue();

                    $fila = $celda->getRow();
                    $columna = $celda->getColumn();

                    if($valorFormateado == "Fecha"){$this->dateC = $columna;}
                    if($valorFormateado == "Ref."){$this->refC = $columna;}
                    if($valorFormateado == "Monto"){$this->amountC = $columna;}
                }
            }
        }
    }

    function positiveAmounts(){
        for ($indiceHoja = 0; $indiceHoja < $this->totalDeHojas; $indiceHoja++) {

            $hojaActual = $this->documento->getSheet($indiceHoja);

            foreach ($hojaActual->getRowIterator() as $fila) {
                foreach ($fila->getCellIterator() as $celda) {

                    $valorFormateado = $celda->getFormattedValue();

                    $fila = $celda->getRow();
                    $columna = $celda->getColumn();

                    $exp = '/^[1-9][\.\d]*(,\d+)?$/';

                    if ($columna == $this->amountC && ($valorFormateado != null || $valorFormateado != "")) {
                        if(preg_match($exp, $valorFormateado)) {
                            array_push($this->amountsArray, $valorFormateado);
                            array_push($this->filasArray, $fila);
                        }
                    }
                }
            }
        }
    }

    function getData(){
        for ($indiceHoja = 0; $indiceHoja < $this->totalDeHojas; $indiceHoja++) {
            $hojaActual = $this->documento->getSheet($indiceHoja);
            $it = 0;
            $it2 = 0;
            foreach ($hojaActual->getRowIterator() as $fila) {
                foreach ($fila->getCellIterator() as $celda) {
                    $valorFormateado = $celda->getFormattedValue();

                    $fila = $celda->getRow();
                    $columna = $celda->getColumn();
                    $exp2 = '/^([0-2][0-9]|3[0-1])(\/|-)(0[1-9]|1[0-2])\2(\d{4})$/';
                    $exp3 = '/^\d+$/';

                    if(count($this->filasArray)-1 >= $it){
                        if ($columna == "A" && $fila == $this->filasArray[$it] &&($valorFormateado != null || $valorFormateado != "")) {
                            if(preg_match($exp2, $valorFormateado)) {
                                array_push($this->DateArray, $valorFormateado);
                                $it++;
                            }
                        }
                    }
                    
                    if(count($this->filasArray)-1 >= $it2){
                        if ($columna == "B" && $fila == $this->filasArray[$it2] &&($valorFormateado != null || $valorFormateado != "")) {
                            if(preg_match($exp3, $valorFormateado)) {
                                array_push($this->RefArray, $valorFormateado);
                                $it2++;
                            }
                        }
                    }
                }
            }
        }
        for ($i=0; $i < count($this->filasArray); $i++) { 
                $movimiento = new stdClass();
                $movimiento->fecha = $this->DateArray[$i];
                $movimiento->ref = $this->RefArray[$i];
                $movimiento->monto = $this->amountsArray[$i];

                array_push($this->movimientos, $movimiento);
            }

            for ($i=0; $i < count($this->movimientos); $i++) { 
                echo $this->movimientos[$i]->fecha, " ", $this->movimientos[$i]->ref, " ", $this->movimientos[$i]->monto, "<br><br>";
            }
    }

    function drotacaExcel(){
        $this->columns();
        $this->positiveAmounts();
        $this->getData();
    }
}