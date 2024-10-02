<?php
require 'vendor/autoload.php'; // Asegúrate de tener PhpSpreadsheet instalado correctamente

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Obtener los datos del formulario
    $nombre = $_POST['nombre'];
    $email = $_POST['email'];
    $telefono = $_POST['telefono'];

    // Ruta del archivo Excel
    $filePath = 'clientes.xlsx';

    // Si el archivo de Excel no existe, crea uno nuevo
    if (!file_exists($filePath)) {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Nombre');
        $sheet->setCellValue('B1', 'Email');
        $sheet->setCellValue('C1', 'Teléfono');
    } else {
        // Cargar el archivo existente
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
    }

    // Encontrar la siguiente fila vacía
    $row = $sheet->getHighestRow() + 1;

    // Escribir los nuevos datos en la siguiente fila vacía
    $sheet->setCellValue("A$row", $nombre);
    $sheet->setCellValue("B$row", $email);
    $sheet->setCellValue("C$row", $telefono);

    // Guardar el archivo de Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);

    // Después de guardar los datos, redirigir al mismo formulario con los campos vacíos
    header("Location: index.html");
    exit;
}
