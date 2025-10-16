<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

header('Content-Type: application/json');

$input = json_decode(file_get_contents('php://input'), true);
if(!$input) { echo json_encode(['ok'=>false,'error'=>'no data']); exit; }

$username = trim($input['username']);
if(strtolower($username) === 'alireza') { echo json_encode(['ok'=>false,'error'=>'excluded']); exit; }

$name = trim($input['name'] ?? '');
$email = trim($input['email'] ?? '');
$version = trim($input['version'] ?? '');
$edition = trim($input['edition'] ?? '');

if(!$username || !$name || !$email || !$version || !$edition) {
    echo json_encode(['ok'=>false,'error'=>'missing fields']); exit;
}

$xlsxFile = __DIR__.'/copy of Almass_2.xlsx';
if(file_exists($xlsxFile)){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($xlsxFile);
    $sheet = $spreadsheet->getActiveSheet();
} else {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->fromArray(['username','name','email','edition','version','saved_at'], null, 'A1');
}

$highestRow = $sheet->getHighestRow()+1;
$sheet->fromArray([$username,$name,$email,$edition,$version,date('Y-m-d H:i:s')], null, 'A'.$highestRow);

$writer = new Xlsx($spreadsheet);
$writer->save($xlsxFile);

echo json_encode(['ok'=>true]);
