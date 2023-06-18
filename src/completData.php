<?php 
 
// Load the database configuration file 
include_once 'dbConfig.php'; 
 
// Include PhpSpreadsheet library autoloader 
require_once '../vendor/autoload.php'; 
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;  
use PhpOffice\PhpSpreadsheet\IOFactory;

if(isset($_POST['importSubmit'])){ 
     
    // Allowed mime types 
    $excelMimes = array('text/xls', 'text/xlsx', 'application/excel', 'application/vnd.msexcel', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); 
     
    // Validate whether selected file is a Excel file 
    if(!empty($_FILES['file']['name']) && in_array($_FILES['file']['type'], $excelMimes)){ 
         
        // If the file is uploaded 
        if(is_uploaded_file($_FILES['file']['tmp_name'])){ 

            // Load the Excel file
            $spreadsheet = IOFactory::load($_FILES['file']['tmp_name']);

            // Select a specific sheet
            $worksheet = $spreadsheet->getActiveSheet();

            // Get the highest row index
            $highestRow = $worksheet->getHighestRow();

            // Iterate through the rows
            for ($row = 2; $row <= $highestRow; $row++) {

                $email = $worksheet->getCell('C' . $row)->getValue();
                $prevQuery = "SELECT * FROM members WHERE email = '".$email."'"; 
                $prevResult = $db->query($prevQuery); 
                $data = $prevResult->fetch_assoc();
                if($prevResult->num_rows > 0){ 

                    $worksheet->setCellValue('A' . $row, $data['first_name']); 
                    $worksheet->setCellValue('B' . $row, $data['last_name']); 

                }else{ 
                     
                }  
            }

            $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
            $writer->save('modified_file.xlsx');
            
             
            $qstring = '?status=succ'; 
        }else{ 
            $qstring = '?status=err'; 
        } 
    }else{ 
        $qstring = '?status=invalid_file'; 
    } 
} 
 
// Redirect to the listing page 
header("Location: result.php".$qstring); 
 
?>