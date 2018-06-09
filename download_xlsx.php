<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

$campos = array('cd_estado','nome_estado','sigla_estado');

/* # USO DO BANCO DE DADOS PARA PEGAR OS DADOS E MONTAR O ARRAY NO FORMATO QUE A FUNÇÃO "fromArray()" NECESSITA
$user = 'root';
$pass = '';
try {
    $db = new PDO('mysql:host=localhost;dbname=adminti', $user, $pass);
} catch (PDOException $e) {
    print "Error!: " . $e->getMessage() . "<br/>";
    die();
}

$result = $db->query('SELECT * from estado');
$rows = $result->fetchAll(PDO::FETCH_ASSOC);

$dados = array();
$posicoes = array();
foreach($rows as $c => $v){
    foreach($campos as $campo){
        $posicoes[] = $v[$campo];
    }
    $dados[] = $posicoes;
    $posicoes = array();
}
array_unshift($dados, $campos); # Adiciona o nome das colunas na primeira linha do excel

$db = null;
*/

# ARRAY MOTANDO DE ACORDO COM O FORMATO QUE A FUNÇÃO "fromArray()" TEM QUE RECEBER
$dados = [
    ['cd_estado','nome_estado','sigla_estado'],
    [1,   'Rio de Janeiro',   'RJ'],
    [2,   'Minas Gerais',   'MG'],
];

#require_once __DIR__ . '/../../src/Bootstrap.php';
#require 'vendor/autoload.php';
require_once 'vendor/phpoffice/phpspreadsheet/src/Bootstrap.php';

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()
    ->fromArray(
        $dados,  // The data to set
        NULL#,        // Array values with this value will not be set
        #'A1'         // Top left coordinate of the worksheet range where
                     //    we want to set these values (default is A1)
    );

// Redirect output to a client’s web browser (Xlsx)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="relatorio.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
exit;
