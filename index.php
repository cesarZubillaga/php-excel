<?php
/**
 * Created by PhpStorm.
 * User: cesar
 * Date: 23/11/16
 * Time: 23:18
 */
include 'vendor/autoload.php';
include 'ExcelCreator.php';
error_reporting(-1);

$creator = new \ExcelCreator();
$singles = array(
    'descripcion', 'titulo', 'contenido'
);

$multiples = array(
    'fotografías' => array(
        'foto 1', 'foto 2', 'foto 3'
    ),
    'usuarios' => array(
        'usuario 1', 'usuario 2', 'usuario 3'
    ),
    'direcciones' => array(
        'direccion 1', 'direccion 2', 'direccion 3'
    )
);
$sheets = array(1, 2, 3, 4, 5, 6);
$excel = new PHPExcel();
foreach ($sheets as $sheet) {
    var_dump($sheet);
    $creator->createSheet($sheet, sprintf("Title %s", $sheet));
    foreach ($singles as $single) {
        $creator->pushSingle(ucfirst($single), sprintf("Value %s", rand(1000, 9999)));
    }
    foreach ($multiples as $key => $values) {
        $creator->pushMultiple(ucfirst($key), $values);
    }
    $creator->setSheet();
}
    $creator->getExcel();
