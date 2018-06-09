<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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

$spreadsheet = new Spreadsheet();

$spreadsheet->getActiveSheet()
    ->fromArray(
        $dados,  // The data to set
        NULL#,        // Array values with this value will not be set
        #'A1'         // Top left coordinate of the worksheet range where
                     //    we want to set these values (default is A1)
    );
$writer = new Xlsx($spreadsheet);
$writer->save('relatorio.xlsx');