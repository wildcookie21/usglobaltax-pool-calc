<?php
if ($_SERVER['REQUEST_METHOD'] !== 'POST' || empty($_FILES['input_sheet'])) {
  http_response_code(400);
  exit('Invalid request.');
}

require_once __DIR__ . '/../bootstrap.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;



if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['input_sheet'])) {
  $fileTmpPath = $_FILES['input_sheet']['tmp_name'];
  $fileName = $_FILES['input_sheet']['name'];
  $fileSize = $_FILES['input_sheet']['size'];
  $fileType = $_FILES['input_sheet']['type'];

  if (preg_match('/INPUT SHEET[^A-Za-z]*([A-Za-z].*?)\.[^.]+$/i', $fileName, $matches)) {
    $clientName = trim($matches[1], " \t\n\r\0\x0B-_");    
  }   

  $allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
  if (!in_array($fileType, $allowedTypes)) {
    throw new RuntimeException('Пожалуйста, загрузите файл Excel (.xlsx или .xls).');
  }

  try {
    $spreadsheet = IOFactory::load($fileTmpPath); // Загружаем и читаем файл
    $sheet = $spreadsheet->getActiveSheet(); // Получаем активный лист
    $highestRow = $sheet->getHighestRow();  // Получаем количество строк

    $actionArr = [];
    $symbolArr = [];
    $quantityArr = [];
    $priceArr = [];
    $totalArr = [];
    $dateGBFormatArr = [];

    for ($row = 2; $row <= $highestRow; $row++) {

      $dateValue = trim($sheet->getCell('A' . $row)->getFormattedValue());

      $dataType = $sheet->getCell('A' . $row)->getDataType();

      if ($dataType == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC) {

        // Преобразуем числовое значение в объект DateTime
        $dateValue = $sheet->getCell('A' . $row)->getValue();
        $dateObject = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($dateValue);

        $formatCode = $sheet->getStyle('A' . $row)
          ->getNumberFormat()
          ->getFormatCode();

        if (stripos($formatCode, 'dd') !== false && stripos($formatCode, 'mm') !== false) {
          $ddPos = stripos($formatCode, 'dd');
          $mmPos = stripos($formatCode, 'mm');

          if ($ddPos < $mmPos) {
            // Формат ячейки DD.MM.YYYY, но данные вводились как MM.DD.YYYY
            $month = $dateObject->format('d');
            $day   = $dateObject->format('m');
            $year  = $dateObject->format('Y');
            $formattedDate = $day . '.' . $month . '.' . $year;
          } elseif ($ddPos > $mmPos) {
            // Формат ячейки MM.DD.YYYY, данные вводились как MM.DD.YYYY
            $month = $dateObject->format('m');
            $day   = $dateObject->format('d');
            $year  = $dateObject->format('Y');
            $formattedDate = $day . '.' . $month . '.' . $year;
          }
        } else {
          require __DIR__ . '/errors/date-format-error.php';
          exit;
        }
      } else {
        // Если это не дата, то просто получаем значение
        $dateValue = $sheet->getCell('A' . $row)->getValue();

        // Если значение выглядит как дата в строковом формате, например, "03/06/2023"
        $pattern = '/(\d{1,2})[\/\.\-\s](\d{1,2})[\/\.\-\s](\d{2,4})/'; // Регулярка для поиска даты

        $replacement = '$2.$1.$3';

        // Преобразуем строку даты в нужный формат
        $formattedDate = preg_replace($pattern, $replacement, $dateValue);

        // Если год двухзначный, добавляем '20' перед годом
        list($day, $month, $year) = explode('.', $formattedDate);
        if (strlen($year) == 2) {
          $year = '20' . $year;  // Преобразуем двухзначный год в четырехзначный
        }

        // Формируем финальную строку в формате 'dd.mm.yyyy'
        $formattedDate = str_pad($day, 2, '0', STR_PAD_LEFT) . '.' . str_pad($month, 2, '0', STR_PAD_LEFT) . '.' . $year;
      }

      // Добавляем отформатированную дату в массив
      $dateGBFormatArr[] = $formattedDate;

      $actionValue = $sheet->getCell('B' . $row)->getValue(); // Получаем Действие (Депозит/Продажа) из столбца B
      $actionArr[] = $actionValue;
      $symbolValue = $sheet->getCell('C' . $row)->getValue(); // Получаем Символ из столбца C
      $symbolArr[] = $symbolValue;
      $quantityValue = $sheet->getCell('D' . $row)->getValue(); // Получаем Количество акций из столбца D
      $quantityArr[] = $quantityValue;
      $priceValue = $sheet->getCell('E' . $row); // Получаем Цену за акцию USD из столбца E
      $priceType = $priceValue->getDataType();
      $totalValue = $sheet->getCell('F' . $row); // Получаем итоговую сумму за все акции из 'D' в USD
      $totalType = $totalValue->getDataType();
      if ($priceType == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA) {
        $calculatedPrice = $priceValue->getCalculatedValue(); // Получаем вычисленное значение если формула
        $priceArr[] = $calculatedPrice;
      } elseif (($priceValue == 0 or $priceValue == '') and !($totalType == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA)) {
        $calculatedTotal = $totalValue->getValue();
        $calculatedPrice = $calculatedTotal / $quantityValue; // если цена за акцию не известна, то итоговую сумму делим кол-во акций
        $priceArr[] = $calculatedPrice;
      } elseif (($priceValue == 0 or $priceValue == '') and ($totalType == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA)) {
        $calculatedTotal = $totalValue->getCalculatedValue();
        $calculatedPrice = $calculatedTotal / $quantityValue; // если цена за акцию не известна, то итоговую сумму делим кол-во акций
        $priceArr[] = $calculatedPrice;
      } else {
        $calculatedPrice = $priceValue->getValue(); // Получаем значение, если нет формулы
        $priceArr[] = $calculatedPrice;
      }

      if ($totalType == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA) {
        $calculatedTotal = $totalValue->getCalculatedValue(); // Получаем вычисленное значение если формула
        $totalArr[] = $calculatedTotal;
      } elseif ($totalValue == 0 or $totalValue == '') {
        $calculatedPrice = $priceValue->getValue();
        $calculatedTotal = $quantityValue * $calculatedPrice; // Получаем значение если значение не проставлено или проставлено 0
        $totalArr[] = $calculatedTotal;
      } else {
        $calculatedTotal = $totalValue->getValue(); // Получаем значение, если нет формулы
        $totalArr[] = $calculatedTotal;
      }

      $quantityValue = '';
      $priceValue = '';
      $totalValue = '';
    }

    $fullTransactionInfoArr = array_map(null, $dateGBFormatArr, $actionArr, $symbolArr, $quantityArr, $priceArr, $totalArr);

    usort($fullTransactionInfoArr, function ($a, $b) {
      $dateA = DateTime::createFromFormat('d.m.Y', $a[0]); // Преобразуем даты из строк в объекты DateTime
      $dateB = DateTime::createFromFormat('d.m.Y', $b[0]); // Преобразуем даты из строк в объекты DateTime

      return $dateA <=> $dateB; // Сортируем по убыванию (по убыванию даты)
    });

    $newSpreadsheet = new Spreadsheet(); //создаём новый exel документ
    $newSheet = $newSpreadsheet->getActiveSheet(); //Получаем активный лист
    $newSheet->setTitle('OUTPUT Transactions');

    $newSheet->getRowDimension(1)->setRowHeight(40); // увеличиваем высоту первой строки
    $newSheet->freezePane('A2'); // Замораживаем первую строку
    $newSheet->getStyle('A1:I1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER); //выравнивание
    $newSheet->getStyle('A1:I1')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);  //выравнивание

    $newSheet->setCellValue('A1', 'Date (GB Format)'); //Редактирование: изменим название столбца A
    $newSheet->setCellValue('B1', 'Action (Deposit/Sale)'); //Редактирование: изменим название столбца B
    $newSheet->setCellValue('C1', 'Symbol'); //Редактирование: изменим название столбца C
    $newSheet->setCellValue('D1', 'Quantity of shares'); //Редактирование: изменим название столбца D
    $newSheet->setCellValue('E1', 'Price per share USD'); //Редактирование: изменим название столбца E
    $newSheet->setCellValue('F1', 'Total USD'); //Редактирование: изменим название столбца F
    $newSheet->setCellValue('G1', 'Spot Exchange Rate (USD/GBP)'); //Редактирование: изменим название столбца G
    $newSheet->setCellValue('H1', 'Price per share GBP'); //Редактирование: изменим название столбца H
    $newSheet->setCellValue('I1', 'Total GBP'); //Редактирование: изменим название столбца I



    $conn = new mysqli(
      $_ENV['DB_HOST'],
      $_ENV['DB_USER'],
      $_ENV['DB_PASS'],
      $_ENV['DB_NAME']
    );

    if ($conn->connect_error) {
      die("Connection failed: " . $conn->connect_error);
    }

    $sql = "SELECT usd_gbp_rate FROM oanda WHERE date = ?";
    $stmt = $conn->prepare($sql);

    foreach ($fullTransactionInfoArr as $fullTransactionInfo) {   // Заполняем данными

      $currentRowOutputTrans = $newSheet->getHighestRow() + 1; // добавляем данные в конец листа      

      $searchDate = $fullTransactionInfo[0];

      $stmt->bind_param("s", $searchDate);
      $stmt->execute(); // Выполнение запроса
      $stmt->bind_result($usdGbpRate); // Получаем результат
      if ($stmt->fetch()) {
        $newSheet->setCellValue('G' . $currentRowOutputTrans, $usdGbpRate); // Если совпадение найдено, выводим курс
        $newSheet->setCellValue('H' . $currentRowOutputTrans, $usdGbpRate * $fullTransactionInfo[4]);
        $newSheet->setCellValue('I' . $currentRowOutputTrans, $usdGbpRate * $fullTransactionInfo[4] * $fullTransactionInfo[3]);
      } else {
        $newSheet->setCellValue('G' . $currentRowOutputTrans, "0"); // Если совпадений нет
        $newSheet->setCellValue('H' . $currentRowOutputTrans, "0");
        $newSheet->setCellValue('I' . $currentRowOutputTrans, "0");
      }


      $actionValueSale = $fullTransactionInfo[1];
      if (strtolower($actionValueSale) === 'sale') {
        $newSheet->getStyle('B' . $currentRowOutputTrans)->getFont()->setBold(true)->setSize(12); // Устанавливаем курсив
      }

      $newSheet->setCellValue('A' . $currentRowOutputTrans, $fullTransactionInfo[0]);
      $newSheet->setCellValue('B' . $currentRowOutputTrans, $actionValueSale);
      $newSheet->setCellValue('C' . $currentRowOutputTrans, $fullTransactionInfo[2]);
      $newSheet->setCellValue('D' . $currentRowOutputTrans, $fullTransactionInfo[3]);
      $newSheet->setCellValue('E' . $currentRowOutputTrans, $fullTransactionInfo[4]);
      $newSheet->setCellValue('F' . $currentRowOutputTrans, $fullTransactionInfo[4] * $fullTransactionInfo[3]);

      $columnsToFormat = ['B', 'C', 'D', 'E', 'F', 'H', 'I'];
      foreach ($columnsToFormat as $column) {
        $newSheet->getStyle($column . '2:' . $column . $newSheet->getHighestRow()) // Применяем формат с двумя знаками после запятой
          ->getNumberFormat()->setFormatCode('#,##0.00'); // Ограничиваем до двух знаков после запятой
      }

      $newSheet->getStyle('A1:I' . $newSheet->getHighestRow()) // Выравниваем все ячейки в столбцах A-I по центру (горизонтально и вертикально)
        ->getAlignment()
        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)  // Горизонтальное выравнивание по центру
        ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);  // Вертикальное выравнивание по центру
    }

    $stmt->close();
    $conn->close();



    $newSheet->getColumnDimension('A')->setWidth(20);  // Дата ширина ячейки
    $newSheet->getColumnDimension('B')->setWidth(25);  // Действие ширина ячейки
    $newSheet->getColumnDimension('C')->setWidth(15);  // Символ ширина ячейки
    $newSheet->getColumnDimension('D')->setWidth(20);  // Количество акций ширина ячейки
    $newSheet->getColumnDimension('E')->setWidth(20);  // Цена за акцию USD ширина ячейки
    $newSheet->getColumnDimension('F')->setWidth(20);  // Итого USD ширина ячейки
    $newSheet->getColumnDimension('G')->setWidth(30);  // Курс обмена USD/GBP ширина ячейки
    $newSheet->getColumnDimension('H')->setWidth(20);  // Цена за акцию GBP ширина ячейки
    $newSheet->getColumnDimension('I')->setWidth(20);  // Итого GBP ширина ячейки


    $highestRow = $newSheet->getHighestRow();  // Получаем количество строк

    $dateArr = [];
    $actionArr = [];
    $symbolArr = [];
    $quantityArr = [];
    $priceUsArr = [];
    $totalUsArr = [];
    $spotArr = [];
    $priceGbArr = [];
    $totalGbArr = [];
    $indexArr = [];
    $index = 0;

    for ($row = 2; $row <= $highestRow; $row++) {
      $dateValue = $newSheet->getCell('A' . $row)->getValue(); // Получаем Дату из столбца А
      $dateArr[] = $dateValue;
      $actionValue = $newSheet->getCell('B' . $row)->getValue(); // Получаем Действие (Депозит/Продажа) из столбца B
      $actionArr[] = $actionValue;
      $symbolValue = $newSheet->getCell('C' . $row)->getValue(); // Получаем Символ из столбца C
      $symbolArr[] = $symbolValue;
      $quantityValue = $newSheet->getCell('D' . $row)->getValue(); // Получаем Количество из столбца D
      $quantityArr[] = $quantityValue;
      $priceUsValue = $newSheet->getCell('E' . $row)->getValue(); // Получаем цену USD из столбца E
      $priceUsArr[] = $priceUsValue;
      $totalUsValue = $newSheet->getCell('F' . $row)->getValue(); // Получаем Сумму USD из столбца F
      $totalUsArr[] = $totalUsValue;
      $spotValue = $newSheet->getCell('G' . $row)->getValue(); // Получаем Курс USD к USD из столбца G
      $spotArr[] = $spotValue;
      $priceGbValue = $newSheet->getCell('H' . $row)->getValue(); // Получаем цену USD к GBP из столбца H
      $priceGbArr[] = $priceGbValue;
      $totalGbValue = $newSheet->getCell('I' . $row)->getValue(); // Получаем Сумму USD к GBP из столбца I
      $totalGbArr[] = $totalGbValue;
      $index++; // порядковый номер для каждой строки
      $indexArr[] = $index;
    }

    $fullInfoArr = array_map(null, $dateArr, $actionArr, $symbolArr, $quantityArr, $priceUsArr, $totalUsArr, $spotArr, $priceGbArr, $totalGbArr, $indexArr);
    $indexValues = array_column($fullInfoArr, 9);  // 9 - это индекс столбца $indexArr, если он последний
    $maxIndex = max($indexValues);   // максимальный индекс

    $saleWordArr = ['sale', 'Sale', 'SALE'];

    $removedIndexes = []; // массив для строк, которые учавствовали в отчёте и выбыли


    $flagSameDayList = true; // флаг чтобы не создать лист несколько раз
    $counterSameDay = 0;

    $flagThirtyDayList = true; // флаг чтобы не создать лист несколько раз
    $counterThirtyDay = 0;

    $flagPoolList = true; // флаг чтобы не создать лист несколько раз
    $counterPool = 0;

    $flagCgtList = true; // флаг чтобы не создать лист несколько раз
    $counterCgt = 0;

    $counterNewTable = 0; // счётчик линии отчёта, отображается в графе SALE

    foreach ($fullInfoArr as $fullInfo) {
      if (in_array($fullInfo[1], $saleWordArr)) {
        $flagSameDayTable = true;
        $flagThirtyDayTable = true;
        $flagPoolTable = true;
        $flagCgtTable = true;
        $totalSumSameDay = 0;
        $totalSharesSameDay = 0;
        $totalSumThirtyDay = 0;
        $totalSharesThirtyDay = 0;
        $totalSumPool = 0;
        $totalSharesPool = 0;
        $lastProcessedItem = 0;
        $counterNewTable++;

        $fullInfoArrCurrent = $fullInfoArr; //из-добавления новых элементов массива и их сортировки в процессе переборки, создаём копию массива
        usort($fullInfoArrCurrent, function ($a, $b) {
          $dateA = DateTime::createFromFormat('d.m.Y', $a[0]); // Преобразуем даты из строк в объекты DateTime
          $dateB = DateTime::createFromFormat('d.m.Y', $b[0]); // Преобразуем даты из строк в объекты DateTime

          return $dateA <=> $dateB; // Сортируем по возрастанию (по возрастанию даты)
        });

        foreach ($fullInfoArrCurrent as $fullInfoForSale) {
          $numberOfSharesSold = $fullInfo[3]; //кол-во проданных акций
          $totalSharesSameDay = 0;
          $totalSharesThirtyDay = 0;
          $totalSharesPool = 0;


          // начало один день
          if ($fullInfo[0] == $fullInfoForSale[0] and !in_array($fullInfoForSale[1], $saleWordArr) and !in_array($fullInfoForSale[9], $removedIndexes)) {

            if ($flagSameDayList) {
              $newList = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($newSpreadsheet, 'OUTPUT Same Day');
              $newSpreadsheet->addSheet($newList);
              $flagSameDayList = false;
            }

            if ($flagSameDayTable) {
              $flagSameDayTable = false;
              $counterSameDay++;
              $newSheet = $newSpreadsheet->setActiveSheetIndexByName('OUTPUT Same Day');
              if ($counterSameDay <= 1) { // предотвращает смещение последущих таблиц "SameDay"
                $currentRowSameDay = $newSheet->getHighestRow();
              }
              $newSheet->setCellValue('A' . $currentRowSameDay, 'SALE');
              $newSheet->setCellValue('B' . $currentRowSameDay, $counterNewTable); //$counterSameDay
              $newSheet->getStyle('A' . $currentRowSameDay . ':B' . $currentRowSameDay)->getFont()->setBold(true);
              $newSheet->getStyle('B' . $currentRowSameDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
              $currentRowSameDay++;
              $newSheet->setCellValue('A' . $currentRowSameDay, 'DATE');
              $newSheet->setCellValue('B' . $currentRowSameDay, $fullInfo[0]);
              $newSheet->getStyle('A' . $currentRowSameDay . ':B' . $currentRowSameDay)->getFont()->setBold(true);
              $newSheet->getStyle('B' . $currentRowSameDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
              $currentRowSameDay++;
              $newSheet->setCellValue('A' . $currentRowSameDay, 'Sale Proceeds total USD');
              $newSheet->setCellValue('B' . $currentRowSameDay, $fullInfo[5]);
              $newSheet->getStyle('A' . $currentRowSameDay . ':B' . $currentRowSameDay)->getFont()->setBold(true);
              $newSheet->getStyle('B' . $currentRowSameDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
              $currentRowSameDay++;
              $newSheet->setCellValue('A' . $currentRowSameDay, 'Spot Excnange Rate @ Date of Sale');
              $newSheet->setCellValue('B' . $currentRowSameDay, $fullInfo[6]);
              $newSheet->getStyle('A' . $currentRowSameDay . ':B' . $currentRowSameDay)->getFont()->setBold(true);
              $newSheet->getStyle('B' . $currentRowSameDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
              $currentRowSameDay++;
              $newSheet->setCellValue('A' . $currentRowSameDay, 'Number of shares sold');
              $newSheet->setCellValue('B' . $currentRowSameDay, $fullInfo[3]);
              $newSheet->getStyle('A' . $currentRowSameDay . ':B' . $currentRowSameDay)->getFont()->setBold(true);
              $newSheet->getStyle('B' . $currentRowSameDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
              $currentRowSameDay++;

              $headers = ['Date', 'Number of shares', 'Cost per share USD', 'Cost total USD', 'Spot Excnange Rate (USD/GBP)', 'Cost total GBP'];
              $row = $currentRowSameDay;
              $col = 'A';
              foreach ($headers as $header) {
                $newSheet->setCellValue($col . $row, $header);
                $col++;
              }

              $firstRowTable = $currentRowSameDay; // начало таблицы

              $newSheet->getColumnDimension('A')->setWidth(35);
              $newSheet->getColumnDimension('B')->setWidth(20);
              $newSheet->getColumnDimension('C')->setWidth(20);
              $newSheet->getColumnDimension('D')->setWidth(20);
              $newSheet->getColumnDimension('E')->setWidth(30);
              $newSheet->getColumnDimension('F')->setWidth(20);
              $newSheet->getStyle('A' . $currentRowSameDay . ':' . 'F' . $currentRowSameDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
              $newSheet->getStyle('A' . $currentRowSameDay . ':' . 'F' . $currentRowSameDay)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
              $currentRowSameDay++;
              $firstDataRow = $currentRowSameDay; //начало таблицы, первая строка с данными(для подсчёта total)
            }

            $newSheet->setCellValue('A' . $currentRowSameDay, $fullInfoForSale[0]);
            $newSheet->setCellValue('B' . $currentRowSameDay, $fullInfoForSale[3]);
            $newSheet->setCellValue('C' . $currentRowSameDay, $fullInfoForSale[4]);
            $newSheet->setCellValue('D' . $currentRowSameDay, $fullInfoForSale[5]);
            $newSheet->setCellValue('E' . $currentRowSameDay, $fullInfoForSale[6]);
            $newSheet->setCellValue('F' . $currentRowSameDay, $fullInfoForSale[8]);

            $removedIndexes[] = $fullInfoForSale[9]; //добавляем индекс строки в массив для строк, которые учавствовали в отчёте и выбыли

            $columnsToFormat = ['B', 'C', 'D', 'F'];
            foreach ($columnsToFormat as $column) {
              $newSheet->getStyle($column . $currentRowSameDay) // Применяем формат с двумя знаками после запятой
                ->getNumberFormat()->setFormatCode('#,##0.00'); // Ограничиваем до двух знаков после запятой
            }

            $currentRowSameDay++;
            $flagForTotalSameDay = true; // для создания total в конце отчёта
            $lastProcessedItem = $fullInfoForSale;
          }
        }

        if ($flagForTotalSameDay) {
          $flagForTotalSameDay = false;
          $totalSumSameDay = 0;
          for ($row = $firstDataRow; $row <= $currentRowSameDay; $row++) {
            $cellValue = $newSheet->getCell('F' . $row)->getValue();
            if (is_numeric($cellValue)) {
              $totalSumSameDay += $cellValue;
            }
          }
          $totalSharesSameDay = 0;
          for ($row = $firstDataRow; $row <= $currentRowSameDay; $row++) {
            $cellValue = $newSheet->getCell('B' . $row)->getValue();
            if (is_numeric($cellValue)) {
              $totalSharesSameDay += $cellValue;
            }
          }
          $newSheet->setCellValue('A' . $currentRowSameDay, 'Total');
          $newSheet->setCellValue('F' . $currentRowSameDay, $totalSumSameDay);
          $newSheet->setCellValue('B' . $currentRowSameDay, $totalSharesSameDay);

          $newSheet->getStyle('A' . $currentRowSameDay)->getFont()->setBold(true);
          $newSheet->getStyle('F' . $currentRowSameDay)->getFont()->setBold(true);
          $newSheet->getStyle('B' . $currentRowSameDay)->getFont()->setBold(true);
          $newSheet->getStyle('F' . $currentRowSameDay)->getNumberFormat()->setFormatCode('#,##0.00');

          $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowSameDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
          $styleArray = [
            'borders' => [
              'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,  // Жирная линия
                'color' => ['argb' => '000000'],  // Черный цвет
              ],
            ],
          ];
          $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowSameDay)->applyFromArray($styleArray);

          $currentRowSameDay++;
          $currentRowSameDay++;

          //начало создания Used up
          $averageSharePriceSameDay = $totalSumSameDay / $totalSharesSameDay; //расчитываем среднюю стоимость акции
          $usedUpSameDay = $totalSharesSameDay; // кол-во акций, которые идут в зачёт текущей продажи
          if ($totalSharesSameDay > $numberOfSharesSold) {
            $usedUpSameDay = $numberOfSharesSold;
          }

          $carriedForward = 0;
          $newSheet->setCellValue('A' . $currentRowSameDay, 'Used up');
          $newSheet->setCellValue('B' . $currentRowSameDay, $usedUpSameDay);
          $newSheet->setCellValue('F' . $currentRowSameDay, $usedUpSameDay * $averageSharePriceSameDay);
          $newSheet->getStyle('F' . $currentRowSameDay)->getNumberFormat()->setFormatCode('#,##0.00');
          if ($totalSharesSameDay > $numberOfSharesSold) { //если итог больше чем кол-во проданных акций, то разницу добавляем в $fullInfoArr
            $carriedForward = $totalSharesSameDay - $usedUpSameDay;
            $maxIndex++;

            $fullInfoArr[] = [$lastProcessedItem[0], $lastProcessedItem[1], $lastProcessedItem[2], $carriedForward, '', '', '', $averageSharePriceSameDay, $averageSharePriceSameDay * $carriedForward, $maxIndex];
          }

          $currentRowSameDay++;
          $newSheet->setCellValue('A' . $currentRowSameDay, 'Carried forward ');
          $newSheet->setCellValue('B' . $currentRowSameDay, $carriedForward);
          $newSheet->setCellValue('F' . $currentRowSameDay, $carriedForward * $averageSharePriceSameDay);
          $newSheet->getStyle('F' . $currentRowSameDay)->getNumberFormat()->setFormatCode('#,##0.00');
          //конец создания Used up

          $currentRowSameDay++;
          $currentRowSameDay++;
        }
        // конец один день

        // начало 30дней
        if ($totalSharesSameDay < $numberOfSharesSold) { //не набрали кол-во продаж в этот же день, запускаем создание месячного отчёта
          $startDate = new DateTime($fullInfo[0]); // Исходная дата
          $endDate = (clone $startDate)->modify('+30 days'); // Дата через 30 дней от startDate
          $startDatePlusOne = (clone $startDate)->modify('+1 day'); // Добавляем 1 день к начальной дате

          foreach ($fullInfoArrCurrent as $fullInfoForSale) {
            $saleDate = new DateTime($fullInfoForSale[0]); // Преобразуем дату для продажи в объект DateTime

            if ($saleDate >= $startDatePlusOne && $saleDate <= $endDate && !in_array($fullInfoForSale[1], $saleWordArr) && !in_array($fullInfoForSale[9], $removedIndexes)) {

              if ($flagThirtyDayList) {
                $newList = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($newSpreadsheet, 'OUTPUT 30 Days');
                $newSpreadsheet->addSheet($newList);
                $flagThirtyDayList = false;
              }

              if ($flagThirtyDayTable) {
                $flagThirtyDayTable = false;
                $counterThirtyDay++;
                $newSheet = $newSpreadsheet->setActiveSheetIndexByName('OUTPUT 30 Days');
                if ($counterThirtyDay <= 1) { // предотвращает смещение последущих таблиц "ThirtyDay"
                  $currentRowThirtyDay = $newSheet->getHighestRow();
                }
                $newSheet->setCellValue('A' . $currentRowThirtyDay, 'SALE');
                $newSheet->setCellValue('B' . $currentRowThirtyDay, $counterNewTable);
                $newSheet->getStyle('A' . $currentRowThirtyDay . ':B' . $currentRowThirtyDay)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowThirtyDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowThirtyDay++;
                $newSheet->setCellValue('A' . $currentRowThirtyDay, 'DATE');
                $newSheet->setCellValue('B' . $currentRowThirtyDay, $fullInfo[0]);
                $newSheet->getStyle('A' . $currentRowThirtyDay . ':B' . $currentRowThirtyDay)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowThirtyDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowThirtyDay++;
                $newSheet->setCellValue('A' . $currentRowThirtyDay, 'Sale Proceeds total USD');
                $newSheet->setCellValue('B' . $currentRowThirtyDay, $fullInfo[5]);
                $newSheet->getStyle('A' . $currentRowThirtyDay . ':B' . $currentRowThirtyDay)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowThirtyDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowThirtyDay++;
                $newSheet->setCellValue('A' . $currentRowThirtyDay, 'Spot Excnange Rate @ Date of Sale');
                $newSheet->setCellValue('B' . $currentRowThirtyDay, $fullInfo[6]);
                $newSheet->getStyle('A' . $currentRowThirtyDay . ':B' . $currentRowThirtyDay)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowThirtyDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowThirtyDay++;
                $newSheet->setCellValue('A' . $currentRowThirtyDay, 'Number of shares sold');
                $newSheet->setCellValue('B' . $currentRowThirtyDay, $fullInfo[3]);
                $newSheet->getStyle('A' . $currentRowThirtyDay . ':B' . $currentRowThirtyDay)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowThirtyDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowThirtyDay++;

                $headers = ['Date', 'Number of shares', 'Cost per share USD', 'Cost total USD', 'Spot Excnange Rate (USD/GBP)', 'Cost total GBP'];
                $row = $currentRowThirtyDay;
                $col = 'A';
                foreach ($headers as $header) {
                  $newSheet->setCellValue($col . $row, $header);
                  $col++;
                }

                $firstRowTable = $currentRowThirtyDay; // начало таблицы

                $newSheet->getColumnDimension('A')->setWidth(35);
                $newSheet->getColumnDimension('B')->setWidth(20);
                $newSheet->getColumnDimension('C')->setWidth(20);
                $newSheet->getColumnDimension('D')->setWidth(20);
                $newSheet->getColumnDimension('E')->setWidth(30);
                $newSheet->getColumnDimension('F')->setWidth(20);
                $newSheet->getStyle('A' . $currentRowThirtyDay . ':' . 'F' . $currentRowThirtyDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $newSheet->getStyle('A' . $currentRowThirtyDay . ':' . 'F' . $currentRowThirtyDay)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $currentRowThirtyDay++;
                $firstDataRow = $currentRowThirtyDay; //начало таблицы, первая строка с данными(для подсчёта total)
              }

              $newSheet->setCellValue('A' . $currentRowThirtyDay, $fullInfoForSale[0]);
              $newSheet->setCellValue('B' . $currentRowThirtyDay, $fullInfoForSale[3]);
              $newSheet->setCellValue('C' . $currentRowThirtyDay, $fullInfoForSale[4]);
              $newSheet->setCellValue('D' . $currentRowThirtyDay, $fullInfoForSale[5]);
              $newSheet->setCellValue('E' . $currentRowThirtyDay, $fullInfoForSale[6]);
              $newSheet->setCellValue('F' . $currentRowThirtyDay, $fullInfoForSale[8]);

              $removedIndexes[] = $fullInfoForSale[9]; //добавляем индекс строки в массив для строк, которые учавствовали в отчёте и выбыли

              $columnsToFormat = ['B', 'C', 'D', 'F'];
              foreach ($columnsToFormat as $column) {
                $newSheet->getStyle($column . $currentRowThirtyDay) // Применяем формат с двумя знаками после запятой
                  ->getNumberFormat()->setFormatCode('#,##0.00'); // Ограничиваем до двух знаков после запятой
              }

              $currentRowThirtyDay++;
              $flagForTotalThirtyDay = true; // для создания total в конце отчёта
              $lastProcessedItem = $fullInfoForSale;
            }
          }

          if ($flagForTotalThirtyDay) {
            $flagForTotalThirtyDay = false;
            $totalSumThirtyDay = 0;
            for ($row = $firstDataRow; $row <= $currentRowThirtyDay; $row++) {
              $cellValue = $newSheet->getCell('F' . $row)->getValue();
              if (is_numeric($cellValue)) {
                $totalSumThirtyDay += $cellValue;
              }
            }
            $totalSharesThirtyDay = 0;
            for ($row = $firstDataRow; $row <= $currentRowThirtyDay; $row++) {
              $cellValue = $newSheet->getCell('B' . $row)->getValue();
              if (is_numeric($cellValue)) {
                $totalSharesThirtyDay += $cellValue;
              }
            }
            $newSheet->setCellValue('A' . $currentRowThirtyDay, 'Total');
            $newSheet->setCellValue('B' . $currentRowThirtyDay, $totalSharesThirtyDay);
            $newSheet->setCellValue('F' . $currentRowThirtyDay, $totalSumThirtyDay);

            $newSheet->getStyle('A' . $currentRowThirtyDay)->getFont()->setBold(true);
            $newSheet->getStyle('F' . $currentRowThirtyDay)->getFont()->setBold(true);
            $newSheet->getStyle('B' . $currentRowThirtyDay)->getFont()->setBold(true);
            $newSheet->getStyle('F' . $currentRowThirtyDay)->getNumberFormat()->setFormatCode('#,##0.00');

            $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowThirtyDay)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $styleArray = [
              'borders' => [
                'allBorders' => [
                  'borderStyle' => Border::BORDER_THIN,  // Жирная линия
                  'color' => ['argb' => '000000'],  // Черный цвет
                ],
              ],
            ];
            $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowThirtyDay)->applyFromArray($styleArray);

            $currentRowThirtyDay++;
            $currentRowThirtyDay++;

            //начало создания Used up

            $averageSharePriceThirtyDay = $totalSumThirtyDay / $totalSharesThirtyDay;
            $usedUpThirtyDay = $totalSharesThirtyDay;
            if (($totalSharesSameDay + $totalSharesThirtyDay) > $numberOfSharesSold) {
              $usedUpThirtyDay = $numberOfSharesSold - $totalSharesSameDay;
            }

            $carriedForward = 0;
            $newSheet->setCellValue('A' . $currentRowThirtyDay, 'Used up');
            $newSheet->setCellValue('B' . $currentRowThirtyDay, $usedUpThirtyDay);
            $newSheet->setCellValue('F' . $currentRowThirtyDay, $usedUpThirtyDay * $averageSharePriceThirtyDay);
            $newSheet->getStyle('F' . $currentRowThirtyDay)->getNumberFormat()->setFormatCode('#,##0.00');
            if ($totalSharesThirtyDay > $numberOfSharesSold) { //если итог больше проданных акций, то разницу добавляем в $fullInfoArr
              $carriedForward = $totalSharesThirtyDay - $usedUpThirtyDay;
              $maxIndex++;

              $fullInfoArr[] = [$lastProcessedItem[0], $lastProcessedItem[1], $lastProcessedItem[2], $carriedForward, '', '', '', $averageSharePriceThirtyDay, $averageSharePriceThirtyDay * $carriedForward, $maxIndex];
            }

            $currentRowThirtyDay++;
            $newSheet->setCellValue('A' . $currentRowThirtyDay, 'Carried forward ');
            $newSheet->setCellValue('B' . $currentRowThirtyDay, $carriedForward);
            $newSheet->setCellValue('F' . $currentRowThirtyDay, $carriedForward * $averageSharePriceThirtyDay);
            $newSheet->getStyle('F' . $currentRowThirtyDay)->getNumberFormat()->setFormatCode('#,##0.00');
            //конец создания Used up

            $currentRowThirtyDay++;
            $currentRowThirtyDay++;
          }
        }
        // конец 30дней


        // начало pool
        if (($totalSharesSameDay + $totalSharesThirtyDay) < $numberOfSharesSold) { //если не набрали в SameDay и ThirtyDay запускаем Pool
          $startDate = new DateTime($fullInfo[0]); // Исходная дата
          $startDateMinusOne = (clone $startDate)->modify('-1 day'); // Отнимает 1 день от начальной дате

          foreach ($fullInfoArrCurrent as $fullInfoForSale) {
            $saleDate = new DateTime($fullInfoForSale[0]); // Преобразуем дату для продажи в объект DateTime

            if ($saleDate <= $startDateMinusOne && !in_array($fullInfoForSale[1], $saleWordArr) && !in_array($fullInfoForSale[9], $removedIndexes)) {

              if ($flagPoolList) {
                $newList = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($newSpreadsheet, 'OUTPUT Pool');
                $newSpreadsheet->addSheet($newList);
                $flagPoolList = false;
              }

              if ($flagPoolTable) {
                $flagPoolTable = false;
                $counterPool++;
                $newSheet = $newSpreadsheet->setActiveSheetIndexByName('OUTPUT Pool');
                if ($counterPool <= 1) { // предотвращает смещение последущих таблиц "Pool"
                  $currentRowPool = $newSheet->getHighestRow();
                }
                $newSheet->setCellValue('A' . $currentRowPool, 'SALE');
                $newSheet->setCellValue('B' . $currentRowPool, $counterNewTable);
                $newSheet->getStyle('A' . $currentRowPool . ':B' . $currentRowPool)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowPool)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowPool++;
                $newSheet->setCellValue('A' . $currentRowPool, 'DATE');
                $newSheet->setCellValue('B' . $currentRowPool, $fullInfo[0]);
                $newSheet->getStyle('A' . $currentRowPool . ':B' . $currentRowPool)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowPool)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowPool++;
                $newSheet->setCellValue('A' . $currentRowPool, 'Sale Proceeds total USD');
                $newSheet->setCellValue('B' . $currentRowPool, $fullInfo[5]);
                $newSheet->getStyle('A' . $currentRowPool . ':B' . $currentRowPool)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowPool)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowPool++;
                $newSheet->setCellValue('A' . $currentRowPool, 'Spot Excnange Rate @ Date of Sale');
                $newSheet->setCellValue('B' . $currentRowPool, $fullInfo[6]);
                $newSheet->getStyle('A' . $currentRowPool . ':B' . $currentRowPool)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowPool)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowPool++;
                $newSheet->setCellValue('A' . $currentRowPool, 'Number of shares sold');
                $newSheet->setCellValue('B' . $currentRowPool, $fullInfo[3]);
                $newSheet->getStyle('A' . $currentRowPool . ':B' . $currentRowPool)->getFont()->setBold(true);
                $newSheet->getStyle('B' . $currentRowPool)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentRowPool++;

                $headers = ['Date', 'Number of shares', 'Cost per share USD', 'Cost total USD', 'Spot Excnange Rate (USD/GBP)', 'Cost total GBP'];
                $row = $currentRowPool;
                $col = 'A';
                foreach ($headers as $header) {
                  $newSheet->setCellValue($col . $row, $header);
                  $col++;
                }

                $firstRowTable = $currentRowPool; // начало таблицы

                $newSheet->getColumnDimension('A')->setWidth(35);
                $newSheet->getColumnDimension('B')->setWidth(20);
                $newSheet->getColumnDimension('C')->setWidth(20);
                $newSheet->getColumnDimension('D')->setWidth(20);
                $newSheet->getColumnDimension('E')->setWidth(30);
                $newSheet->getColumnDimension('F')->setWidth(20);
                $newSheet->getStyle('A' . $currentRowPool . ':' . 'F' . $currentRowPool)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $newSheet->getStyle('A' . $currentRowPool . ':' . 'F' . $currentRowPool)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $currentRowPool++;
                $firstDataRow = $currentRowPool; //начало таблицы, первая строка с данными(для подсчёта total)
              }

              $newSheet->setCellValue('A' . $currentRowPool, $fullInfoForSale[0]);
              $newSheet->setCellValue('B' . $currentRowPool, $fullInfoForSale[3]);
              $newSheet->setCellValue('C' . $currentRowPool, $fullInfoForSale[4]);
              $newSheet->setCellValue('D' . $currentRowPool, $fullInfoForSale[5]);
              $newSheet->setCellValue('E' . $currentRowPool, $fullInfoForSale[6]);
              $newSheet->setCellValue('F' . $currentRowPool, $fullInfoForSale[8]);

              $removedIndexes[] = $fullInfoForSale[9]; //добавляем индекс строки в массив для строк, которые учавствовали в отчёте и выбыли

              $columnsToFormat = ['B', 'C', 'D', 'F'];
              foreach ($columnsToFormat as $column) {
                $newSheet->getStyle($column . $currentRowPool) // Применяем формат с двумя знаками после запятой
                  ->getNumberFormat()->setFormatCode('#,##0.00'); // Ограничиваем до двух знаков после запятой
              }

              $currentRowPool++;
              $flagForTotalPool = true; // для создания total в конце отчёта
              $lastProcessedItem = $fullInfoForSale;
            }
          }

          if ($flagForTotalPool) {
            $flagForTotalPool = false;
            $totalSumPool = 0;
            for ($row = $firstDataRow; $row <= $currentRowPool; $row++) {
              $cellValue = $newSheet->getCell('F' . $row)->getValue();
              if (is_numeric($cellValue)) {
                $totalSumPool += $cellValue;
              }
            }
            $totalSharesPool = 0;
            for ($row = $firstDataRow; $row <= $currentRowPool; $row++) {
              $cellValue = $newSheet->getCell('B' . $row)->getValue();
              if (is_numeric($cellValue)) {
                $totalSharesPool += $cellValue;
              }
            }
            $newSheet->setCellValue('A' . $currentRowPool, 'Total');
            $newSheet->setCellValue('B' . $currentRowPool, $totalSharesPool);
            $newSheet->setCellValue('F' . $currentRowPool, $totalSumPool);

            $newSheet->getStyle('A' . $currentRowPool)->getFont()->setBold(true);
            $newSheet->getStyle('F' . $currentRowPool)->getFont()->setBold(true);
            $newSheet->getStyle('B' . $currentRowPool)->getFont()->setBold(true);
            $newSheet->getStyle('F' . $currentRowPool)->getNumberFormat()->setFormatCode('#,##0.00');

            $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowPool)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $styleArray = [
              'borders' => [
                'allBorders' => [
                  'borderStyle' => Border::BORDER_THIN,  // Жирная линия
                  'color' => ['argb' => '000000'],  // Черный цвет
                ],
              ],
            ];
            $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowPool)->applyFromArray($styleArray);

            $currentRowPool++;
            $currentRowPool++;

            //начало создания Used up
            $averageSharePricePool = $totalSumPool / $totalSharesPool;
            $usedUpPool = $totalSharesPool;
            if (($totalSharesSameDay + $totalSharesThirtyDay + $totalSharesPool) > $numberOfSharesSold) {

              $usedUpPool = $numberOfSharesSold - ($totalSharesSameDay + $totalSharesThirtyDay);
            }

            $carriedForward = 0;
            $newSheet->setCellValue('A' . $currentRowPool, 'Used up');
            $newSheet->setCellValue('B' . $currentRowPool, $usedUpPool);
            $newSheet->setCellValue('F' . $currentRowPool, $usedUpPool * $averageSharePricePool);
            $newSheet->getStyle('F' . $currentRowPool)->getNumberFormat()->setFormatCode('#,##0.00');
            if ($totalSharesPool > $numberOfSharesSold) { //если итог больше проданных акций, то разницу добавляем в $fullInfoArr
              $carriedForward = $totalSharesPool - $usedUpPool;
              $maxIndex++;

              $fullInfoArr[] = [$lastProcessedItem[0], $lastProcessedItem[1], $lastProcessedItem[2], $carriedForward, '', '', '', $averageSharePricePool, $averageSharePricePool * $carriedForward, $maxIndex];
            }

            $currentRowPool++;
            $newSheet->setCellValue('A' . $currentRowPool, 'Carried forward ');
            $newSheet->setCellValue('B' . $currentRowPool, $carriedForward);
            $newSheet->setCellValue('F' . $currentRowPool, $carriedForward * $averageSharePricePool);
            $newSheet->getStyle('F' . $currentRowPool)->getNumberFormat()->setFormatCode('#,##0.00');

            $currentRowPool++;
            $currentRowPool++;
            //конец создания Used up
          }
          //конец pool
        }


        // начало CGT Culc
        if ($flagCgtList) {
          $newList = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($newSpreadsheet, 'OUTPUT CGT Culc');
          $newSpreadsheet->addSheet($newList);
          $flagCgtList = false;
        }

        if ($flagCgtTable) {
          $flagCgtTable = false;
          $counterCgt++;
          $newSheet = $newSpreadsheet->setActiveSheetIndexByName('OUTPUT CGT Culc');
          if ($counterCgt <= 1) { // предотвращает смещение последущих таблиц "CGT"
            $currentRowCgt = $newSheet->getHighestRow();
            $newSheet->setCellValue('A' . $currentRowCgt, 'TOTAL GAIN');
            $positionTotalTotalGain = $currentRowCgt;
            $newSheet->getStyle('A' . $currentRowCgt)->getFont()->setBold(true);
            $currentRowCgt++;
            $newSheet->setCellValue('A' . $currentRowCgt, 'TOTAL LOSS');
            $positionTotalTotalLoss = $currentRowCgt;
            $newSheet->getStyle('A' . $currentRowCgt)->getFont()->setBold(true);
            $currentRowCgt++;
            $currentRowCgt++;
          }

          $newSheet->setCellValue('A' . $currentRowCgt, 'SALE');
          $newSheet->setCellValue('B' . $currentRowCgt, $counterNewTable);
          $newSheet->getStyle('A' . $currentRowCgt . ':B' . $currentRowCgt)->getFont()->setBold(true);
          $newSheet->getStyle('B' . $currentRowCgt)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
          $currentRowCgt++;
          $newSheet->setCellValue('A' . $currentRowCgt, 'DATE');
          $newSheet->setCellValue('B' . $currentRowCgt, $fullInfo[0]);
          $newSheet->getStyle('A' . $currentRowCgt . ':B' . $currentRowCgt)->getFont()->setBold(true);
          $newSheet->getStyle('B' . $currentRowCgt)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
          $currentRowCgt++;
          $newSheet->setCellValue('A' . $currentRowCgt, 'Sale Proceeds total USD');
          $newSheet->setCellValue('B' . $currentRowCgt, $fullInfo[5]);
          $newSheet->getStyle('A' . $currentRowCgt . ':B' . $currentRowCgt)->getFont()->setBold(true);
          $newSheet->getStyle('B' . $currentRowCgt)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
          $currentRowCgt++;
          $newSheet->setCellValue('A' . $currentRowCgt, 'Spot Excnange Rate @ Date of Sale');
          $newSheet->setCellValue('B' . $currentRowCgt, $fullInfo[6]);
          $newSheet->getStyle('A' . $currentRowCgt . ':B' . $currentRowCgt)->getFont()->setBold(true);
          $newSheet->getStyle('B' . $currentRowCgt)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
          $currentRowCgt++;
          $newSheet->setCellValue('A' . $currentRowCgt, 'Number of shares sold');
          $newSheet->setCellValue('B' . $currentRowCgt, $fullInfo[3]);
          $newSheet->getStyle('A' . $currentRowCgt . ':B' . $currentRowCgt)->getFont()->setBold(true);
          $newSheet->getStyle('B' . $currentRowCgt)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
          $currentRowCgt++;

          $headers = ['Details', 'Sale Proceeds GPB', 'Cost - Same Day GBP', 'Cost - Next 30 Days GBP', 'Cost - POOL GBP', 'Number of shares'];
          $row = $currentRowCgt;
          $col = 'A';
          foreach ($headers as $header) {
            $newSheet->setCellValue($col . $row, $header);
            $col++;
          }

          $firstRowTable = $currentRowCgt; // начало таблицы

          $newSheet->getColumnDimension('A')->setWidth(35);
          $newSheet->getColumnDimension('B')->setWidth(20);
          $newSheet->getColumnDimension('C')->setWidth(20);
          $newSheet->getColumnDimension('D')->setWidth(20);
          $newSheet->getColumnDimension('E')->setWidth(30);
          $newSheet->getColumnDimension('F')->setWidth(20);
          $newSheet->getStyle('A' . $currentRowCgt . ':' . 'F' . $currentRowCgt)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
          $newSheet->getStyle('A' . $currentRowCgt . ':' . 'F' . $currentRowCgt)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
          $currentRowCgt++;
        }

        $newSheet->setCellValue('A' . $currentRowCgt, 'SALE');
        $newSheet->setCellValue('B' . $currentRowCgt, $fullInfo[8]);
        $newSheet->getStyle('B' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'Number of shares Sold');
        $newSheet->setCellValue('F' . $currentRowCgt, $fullInfo[3]);
        $newSheet->getStyle('F' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'COST');
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'Number of shares Same Day');
        $newSheet->setCellValue('C' . $currentRowCgt, $usedUpSameDay * $averageSharePriceSameDay);
        $newSheet->setCellValue('F' . $currentRowCgt, $usedUpSameDay);
        $newSheet->getStyle('C' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $newSheet->getStyle('F' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'Number of shares Next 30 Days');
        $newSheet->setCellValue('D' . $currentRowCgt, $usedUpThirtyDay * $averageSharePriceThirtyDay);
        $newSheet->setCellValue('F' . $currentRowCgt, $usedUpThirtyDay);
        $newSheet->getStyle('D' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $newSheet->getStyle('F' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'Number of shares POOL');
        $newSheet->setCellValue('E' . $currentRowCgt, $usedUpPool * $averageSharePricePool);
        $newSheet->setCellValue('F' . $currentRowCgt, $usedUpPool);
        $newSheet->getStyle('E' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $newSheet->getStyle('F' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');


        $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowCgt)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $styleArray = [
          'borders' => [
            'allBorders' => [
              'borderStyle' => Border::BORDER_THIN,  // Жирная линия
              'color' => ['argb' => '000000'],  // Черный цвет
            ],
          ],
        ];
        $newSheet->getStyle('A' . $firstRowTable . ':' . 'F' . $currentRowCgt)->applyFromArray($styleArray);

        $currentRowCgt++;
        $currentRowCgt++;

        $newSheet->setCellValue('A' . $currentRowCgt, 'TOTAL SALE');
        $totalSale = $fullInfo[8];
        $newSheet->setCellValue('B' . $currentRowCgt, $totalSale);
        $newSheet->getStyle('B' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'TOTAL COST');
        $totalCost = ($usedUpSameDay * $averageSharePriceSameDay) + ($usedUpThirtyDay * $averageSharePriceThirtyDay) + ($usedUpPool * $averageSharePricePool);
        $newSheet->setCellValue('B' . $currentRowCgt, $totalCost);
        $newSheet->getStyle('B' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'TOTAL GAIN');
        $newSheet->getStyle('A' . $currentRowCgt)->getFont()->setBold(true);
        if ($totalSale > $totalCost) {
          $totalGain = $totalSale - $totalCost;
        } elseif ($totalSale <= $totalCost) {
          $totalGain = 0;
        }
        $newSheet->setCellValue('B' . $currentRowCgt, $totalGain);
        $newSheet->getStyle('B' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $newSheet->getStyle('B' . $currentRowCgt)->getFont()->setBold(true);
        $currentRowCgt++;
        $newSheet->setCellValue('A' . $currentRowCgt, 'TOTAL LOSS');
        $newSheet->getStyle('A' . $currentRowCgt)->getFont()->setBold(true);
        if ($totalGain > 0) {
          $totalLoss = 0;
        } else {
          $totalLoss = $totalCost - $totalSale;
        }
        $newSheet->setCellValue('B' . $currentRowCgt, $totalLoss);
        $newSheet->getStyle('B' . $currentRowCgt)->getNumberFormat()->setFormatCode('#,##0.00');
        $newSheet->getStyle('B' . $currentRowCgt)->getFont()->setBold(true);
        $totalTotalGain += $totalGain;
        $totalTotalLoss += $totalLoss;

        $usedUpPool = 0;
        $usedUpSameDay = 0;
        $usedUpThirtyDay = 0;
        $averageSharePriceSameDay = 0;
        $averageSharePriceThirtyDay = 0;
        $averageSharePricePool = 0;
        $totalGain = 0;
        $totalCost = 0;
        $totalSale = 0;

        $currentRowCgt++;
        $currentRowCgt++;
        // конец CGT Culc

      }
    }

    $newSheet = $newSpreadsheet->setActiveSheetIndexByName('OUTPUT CGT Culc');
    $newSheet->setCellValue('B' . $positionTotalTotalGain, $totalTotalGain);
    $newSheet->getStyle('B' . $positionTotalTotalGain)->getFont()->setBold(true);
    $newSheet->getStyle('B' . $positionTotalTotalGain)->getNumberFormat()->setFormatCode('#,##0.00');
    $newSheet->setCellValue('B' . $positionTotalTotalLoss, $totalTotalLoss);
    $newSheet->getStyle('B' . $positionTotalTotalLoss)->getFont()->setBold(true);
    $newSheet->getStyle('B' . $positionTotalTotalLoss)->getNumberFormat()->setFormatCode('#,##0.00');


    $writer = new Xlsx($newSpreadsheet); // Создаем объект для записи в файл
    $outputFileName = 'CGT' . ' ' . $clientName . ' ' . date('d.m.Y') . '.xlsx'; // Генерируем имя для сохранения файла

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . $outputFileName . '"');
    header('Cache-Control: no-cache, no-store, must-revalidate');
    header('Pragma: no-cache');
    header('Expires: 0');

    $writer->save('php://output');
    exit;
  } catch (Throwable $e) {
    echo 'Ошибка при обработке файла: ', $e->getMessage();
  }
} else {
  echo "Пожалуйста, выберите файл для загрузки.";
}
