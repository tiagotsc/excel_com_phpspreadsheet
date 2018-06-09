<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
#use PhpOffice\PhpSpreadsheet\Style\Border;
#use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

$campos = array('cd_estado','nome_estado','sigla_estado');

/* # USO DO BANCO DE DADOS PARA PEGAR OS DADOS
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
foreach($rows as $c => $v){
    foreach($campos as $campo){
        $posicoes[] = $v[$campo];
    }
    $dados[] = $posicoes;
    $posicoes = array();
}
$db = null;
*/

# ARRAY MONTADO NO MESMO FORMATO DO BANCO DE DADOS
$rows = array(
            array(
                'cd_estado' => 1,
                'nome_estado' => 'Rio de Janeiro',
                'sigla_estado' => 'RJ'
            ),
            array(
                'cd_estado' => 2,
                'nome_estado' => 'Minas Gerais',
                'sigla_estado' => 'MG'
            )
        );

$dados = array();
$posicoes = array();

require_once 'vendor/phpoffice/phpspreadsheet/src/Bootstrap.php';

$spreadsheet = new Spreadsheet();

$contCampo = 1;
# Estiliza a primeira coluna em negrito
#$spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);

foreach($campos as $campo){ 
            
    # Estiliza a coluna em negrito
    $spreadsheet->getActiveSheet()->getStyle($contCampo)->getFont()->setBold(true);
    # Cria a coluna
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($contCampo, 1,  $campo);
    # Próxima coluna
    $contCampo++;
}

# Conteúdo a partir da segunda linha
$linha = 2;

if(count($rows) > 0){
    
    # Controla o campo da linha
    $coluna = 1;
    foreach($rows as $valor){ # Alimenta as colunas com o conteúdo
        
        # Remove o negrito
        $spreadsheet->getActiveSheet()->getStyle('A2')->getFont()->setBold(false);
        
        # Remove o negrito
        $spreadsheet->getActiveSheet()->getStyle('B2')->getFont()->setBold(false);
        foreach($campos as $campo){
        
            # Estiliza removendo o negrito
            $spreadsheet->getActiveSheet()->getStyle($coluna)->getFont()->setBold(false);
            
                # Exemplos: 10/06/2015 | 10/06/15 | 2015-06-10 | 15-06-10
                if(preg_match('/^[0-9]{2,4}(-|\/)[0-9]{2}(-|\/)[0-9]{2,4}$/', trim($valor[$campo]))){ # Verifica se o conteúdo é uma data
                    # Adiciona o conteúdo - Data no formato Excel
                    # Se for data converte a string data para o formato data do Excel
                    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($coluna, $linha,  Date::PHPToExcel(utf8_encode(trim($valor[$campo])))); #echo trim($valor[$campo]); echo '<br>'; echo $ExcelDateValue; exit();
                    $spreadsheet->getActiveSheet()->getStyleByColumnAndRow($coluna, $linha)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); #PHPExcel_Style_NumberFormat::FORMAT_DATE_DMYSLASH
                
                }elseif(is_numeric(trim($valor[$campo])) and strlen(trim($valor[$campo])) >= 12){ # Se for um número muito grande converte pra string
                    
                    $type = DataType::TYPE_STRING;
                    #$type = DataType::TYPE_NUMERIC;
                    $spreadsheet->getActiveSheet()->getCellByColumnAndRow($coluna, $linha)->setValueExplicit(trim($valor[$campo]), $type);
                    
                }else{ # Mantem o formato padrão
                    # Adiciona o conteúdo
                    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($coluna, $linha,  utf8_encode(trim($valor[$campo])));
                    
                }

            # Segue pra próxima coluna
            $coluna++;
        }
    
        # Retorno a primeira coluna    
        $coluna = 1;
    
        # Avança pra próxima linha
        $linha++;    
    }
}

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
