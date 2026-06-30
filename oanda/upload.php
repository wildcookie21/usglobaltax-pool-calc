<?php
if ($_SERVER['REQUEST_METHOD'] !== 'POST' || empty($_FILES['oanda_excel'])) {
  http_response_code(400);
  exit('Invalid request.');
}

require_once __DIR__ . '/../bootstrap.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['oanda_excel'])) {  
  $fileTmpPath = $_FILES['oanda_excel']['tmp_name'];
  $fileName = $_FILES['oanda_excel']['name'];
  $fileSize = $_FILES['oanda_excel']['size'];
  $fileType = $_FILES['oanda_excel']['type'];
  
  $allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
  if (!in_array($fileType, $allowedTypes)) {
    echo "Ошибка: Пожалуйста, загрузите файл Excel (.xlsx или .xls).";
    exit;
  }

  try {    
    $spreadsheet = IOFactory::load($fileTmpPath);  
    $sheet = $spreadsheet->getActiveSheet(); 
    $highestRow = $sheet->getHighestRow();

    $conn = new mysqli(
      $_ENV['DB_HOST'],
      $_ENV['DB_USER'],
      $_ENV['DB_PASS'],
      $_ENV['DB_NAME']
    );

    if ($conn->connect_error) {
      die("Connection failed: " . $conn->connect_error);
    }

    $rows = $sheet->toArray();
    foreach ($rows as $row) {
      $date = $row[0];      
      $formattedDate = preg_replace_callback(
        '/(\d{2})\/(\d{2})\/(\d{2})/',
        function ($matches) {          
          return $matches[2] . '.' . $matches[1] . '.20' . $matches[3];
        },
        $date
      );
      $rate = str_replace(',', '.', $row[1]);
      
      $stmt = $conn->prepare("INSERT INTO oanda (date, usd_gbp_rate) VALUES (?, ?)");
      $stmt->bind_param("sd", $formattedDate, $rate);

      if ($stmt->execute()) {
        echo "Данные для $formattedDate успешно загружены!\n";
      } else {
        echo "Ошибка при загрузке данных для $formattedDate: " . $stmt->error . "\n";
      }
    }    
    $conn->close();

  } catch (Exception $e) {
    echo 'Ошибка при обработке файла: ', $e->getMessage();
  }
} else {
  echo "Пожалуйста, выберите файл для загрузки.";
}