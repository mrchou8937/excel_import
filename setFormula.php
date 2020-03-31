//phpexcel 在zentao中案例
$objExcel = $this->app->loadClass('PHPExcel');

// 創建文件格式寫入對象實例, uncomment
$objWriter = new PHPExcel_Writer_Excel5($objExcel);     // 用於其他版本格式

$objExcel->setActiveSheetIndex(0);
$objActSheet = $objExcel->getActiveSheet();

//設置A1單元格的選擇列表
$objValidation = $objActSheet->getCell("A1")->getDataValidation();
$objValidation -> setType(PHPExcel_Cell_DataValidation::TYPE_LIST)
    -> setErrorStyle(PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
    -> setAllowBlank(false)
    -> setShowInputMessage(true)
    -> setShowErrorMessage(true)
    -> setShowDropDown(true)
    -> setErrorTitle('輸入的值有誤')
    -> setError('您輸入的值不在下拉框列表內.')
    -> setPromptTitle('設備類型')
    //方法一
    //把sheet名为mySheet2的A1,A2,A3作为选项
		// -> setFormula1('mySheet2!$A$1:$A$3');
		
		//方法二
    //设置为具体的内容
		-> setFormula1('"方案1,方案2,方案3"');
		
		//方法三
    //设置为变量内容,例如:$myStr = 'select1,select2,select3'
  	// -> setFormula1('"'.$myStr.'"');

//設置單元格顏色
$objStyleA1 = $objActSheet ->getStyle('A1');
$objStyleA1 ->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
//設置CELL填充顏色
$objFillA1 = $objStyleA1->getFill();
$objFillA1->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objFillA1->getStartColor()->setARGB('FFcdcdff');

//設置當前活動sheet的名稱
$objActSheet->setTitle('mySheet1');



//==========================================================//
// //配合	方法一
// //新增一个sheet，命名为mySheet2
// $objExcel->createSheet();


// // Add some data to the second sheet, resembling some different data types

// //指定開始改第二個工作表
// $objExcel->setActiveSheetIndex(1);
// $objExcel->getActiveSheet()->setCellValue('K1', 'item1');
// $objExcel->getActiveSheet()->setCellValue('K2', 'item2');
// $objExcel->getActiveSheet()->setCellValue('K3', 'item3');
// $objExcel->getActiveSheet()->setTitle('mySheet2');

// //改完第二工作表，指定第一工作表繼續編輯
// $objExcel->setActiveSheetIndex(0);
//==========================================================//



$outputFileName = "output.xls";

header("Content-Type: application/force-download");
header("Content-Type: application/octet-stream");
header("Content-Type: application/download");
header('Content-Disposition:inline;filename="'.$outputFileName.'"');
header("Content-Transfer-Encoding: binary");
header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
header("Pragma: no-cache");
$objWriter->save('php://output');
