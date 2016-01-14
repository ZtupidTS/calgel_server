<?php
//=====================================================================================================
//этот файл точка входа - в него посылаются данные виджетом через POST запрос
//алгоритм работы - в зависимости от значения POST["ftype"]
//файл состоит из следующих встроенных фунций
// updateValue - содержит алгоритм и бизнес логику обработки docx
// recursive_remove_directory - удаление директорий рекурсивно
// change_docx - основная функция обработки docx
// алгоритм работы
// 1.В зависимости от типа POST['ftype'] определяются параметры входных и выходных файлов
// 2.В зависимости от типа POST['ftype'] определяется ветка алгоритма (обработка docx или xlsx файл)
// 3.1 Обработка docx
// 3.1.1 Запускается change_docx, результат ее работы упаковывается в zip архив
// 3.1.1.1 Обработка внутри  change_docx
// 3.1.1.2 Создается папка temp либо очищается рекурсивно с использованием recursive_remove_directory
// 3.1.1.3 В temp распаковывается через unzip содержимоей файла шаблона
// 3.1.1.4 Содержимое файла temp/word/document.xml читается в строку и передается в updateValue -  где обрабатывается с учетом бизнес логики
// 3.1.1.5 Файл temp/word/document.xml удаляется и создается заново, в новый файл записывается результат из updateValue
// 3.1.1.6 содержимое каталога temp запаковывается в zip, работа change_docx закончена
// 3.1.2 zip архив из шага выше помещается в другой архив вместе с периеменованием zip файла в docx
// 3.2 Обработка xlsx
// 3.2.1 Данные сбрасываются в Excel по циклу
// 4.Дальнейшая обработка не описана т.к. не разбирался, там идет выгрузка в гугл драйв и может чтото еще....
//=====================================================================================================
//@file_put_contents("now-wc.txt","point1");
if (empty($_POST["ftype"])){
    exit();
}
//@file_put_contents("now-wc.txt","post:".implode("#",$_POST));
$ftype = $_POST["ftype"];
$dognum = $_POST["ddnumber"];

//1.=========================================================
switch ($ftype) {
    case 'contract':
        $template_file_name = "dog-calgelunion.docx";
        $out_file_name = "contract_".date("Ymd_His")."_N_". $dognum.".docx";
        $valueToCheck = "Договор № %ddnumber% от %nowdate%";
        break;
    case 'invoice':
        $template_file_name = "invoice-calgelunion.docx";
        $out_file_name = "invoice_".date("Ymd_His")."_N_". $dognum.".docx";
        $valueToCheck = "Счет № %ddnumber% от %nowdate%";
        break;
	case 'act':
        $template_file_name = "act-calgelunion.docx";
        $out_file_name = "act_".date("Ymd_His")."_N_". $dognum.".docx";
        $valueToCheck = "Акт № %ddnumber% от %nowdate%";		
        break;	
	case 'tradeinvoice':
        $template_file_name = "invoice1.xlsx";
        //$out_file_name = "tradeinvoice_".date("Ymd_His")."_N_". $dognum.".xlsx";
		$out_file_name = "tradeinvoice_".date("Ymd_His").".xlsx";
        $valueToCheck = "Счет № %ddnumber% от %nowdate%";
		@file_put_contents("now-wc.txt","tradeinvoice post:");
        break;	
	case 'tradeact':
        $template_file_name = "tradeact1.xlsx";
        //$out_file_name = "tradeinvoice_".date("Ymd_His")."_N_". $dognum.".xlsx";
		$out_file_name = "tradeact_".date("Ymd_His").".xlsx";
        $valueToCheck = "Счет № %ddnumber% от %nowdate%";
		@file_put_contents("now-wc.txt","tradeinvoice post:");
        break;		
}

//error_reporting(E_ALL);
//ini_set('display_errors', 'On');
require_once dirname(__FILE__) . '/num_to_str.php';

$base_path = "http://calgel.ru/esp/files/";

function updateValue($val, $data){
    //заменяет в $val значения на $data

    $arrToSearch = array(
        "%ddnumber%",
        "%daynum%",
        "%monthnum%",
        "%yearnum%",
        "%ddate%",
        "%zfullname%",
        "%wdays%",
        "%zprice%",
        "%zpricetext%",
        "%requisite%"
    );

    if (isset($data["zpricetext"])) {
        $ns = new Num_to_str;
        $strval = $data["zpricetext"];
        $data["zpricetext"] = $ns->get($strval);
    }
    $arrToSet = array(
        $data["ddnumber"],
        $data["daynum"],
        $data["monthnum"],
        $data["yearnum"],
        $data["ddate"],
        $data["zfullname"],
        $data["wdays"],
        $data["zprice"],
        $data["zpricetext"],
        $data["requisite"]
    );

    $res = str_replace($arrToSearch, $arrToSet, $val);
    return $res;
}


function recursive_remove_directory($directory, $empty = false) {
		if (substr($directory, -1) == '/') {
			$directory = substr($directory, 0, -1);
		}	
		if (!file_exists($directory) || !is_dir($directory)) {
			return false;
		} elseif (is_readable($directory)) {
			$handle = opendir($directory);			
			while (false !== ($item = readdir($handle))) {
				if ($item != '.' && $item != '..') {
					$path = $directory.'/'.$item;
					if (is_dir($path)) {
						recursive_remove_directory($path);
					} else {
						unlink($path);
					}
				}
			}
			closedir($handle);
			if ($empty == false) {
				if (!rmdir($directory)) {
					return false;
				}
			}
	
		}
		return true;
	}
function change_xlsx($file, $docname, $leadData, $ddtype) {
	//@file_put_contents("now-wc.txt","point2:change_xlsx:".$docname,FILE_APPEND);
	if (is_file('files/'.$docname)) unlink('files/'.$docname);
	//if (!copy($file, 'files/'.$docname)) { } else {
	if (!copy($file, $docname)) { } else {	
		//продолжаем обработку
		
		
		//работа с xlsx		
		ini_set('include_path', ini_get('include_path').';../PHPExcel-1.8/Classes/');
		require_once './PHPExcel-1.8/Classes/PHPExcel.php';
		require_once './PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';
		require_once './PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';
		require_once './PHPExcel-1.8/Classes/PHPExcel/Cell/AdvancedValueBinder.php';
		require_once './PHPExcel-1.8/Classes/PHPExcel/Writer/IWriter.php';
		

		$objPHPExcel = new PHPExcel();
		$objReader = PHPExcel_IOFactory::createReader('Excel2007');		
		$filenamestr = $docname;
		//$filenamestr = "invoice1.xlsx";
		$objPHPExcel = $objReader->load($filenamestr);
		PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );
		
		// Add some data
		//insert row
		//$objPHPExcel->getActiveSheet()->insertNewRowBefore(7, 2);
		//объединить ячейки
		//$objPHPExcel->getActiveSheet()->mergeCells('A18:E22');

		$objPHPExcel->setActiveSheetIndex(0);
		$dname = "";
		if ($ddtype=='tradeact') {
			$dname='Акт';
		} else {
			$dname='Счет';
		}
		$strB11val= $dname." №".$leadData['ddnumber'].' от '.$leadData['ddate'];
		$objPHPExcel->getActiveSheet()->SetCellValue('B11', $strB11val);
		$objPHPExcel->getActiveSheet()->SetCellValue('G17', "".$leadData['recieve']);
		$objPHPExcel->getActiveSheet()->SetCellValue('G19', "".$leadData['recieve']);
		
		

		$mainarr = explode("!-!",$leadData['alltoinvioce']);
		$i=0;
		$inum = 22;
		foreach($mainarr as $elementarr) {
			if(("".$elementarr)==="") {
				
			} else {
				$inum = 22 + $i;
				$objPHPExcel->getActiveSheet()->insertNewRowBefore($inum, 1); //1 row before22
				$strB = 'B'.$inum;
				$strBVal = "".($i+1);
				$strC = 'C'.$inum;
				$strCVal = "".$elementarr;
				$strW = 'W'.$inum;				
				$strX = 'X'.$inum;
				$strY = 'Y'.$inum;
				$strZ = 'Z'.$inum;
				$strAA = 'AA'.$inum;
				$objPHPExcel->getActiveSheet()->SetCellValue($strB, $strBVal);
				$objPHPExcel->getActiveSheet()->SetCellValue($strC, $strCVal);
				$objPHPExcel->getActiveSheet()->SetCellValue($strX, '');
				$objPHPExcel->getActiveSheet()->SetCellValue($strY, '');
				$objPHPExcel->getActiveSheet()->SetCellValue($strZ, '');
				$objPHPExcel->getActiveSheet()->SetCellValue($strAA, '');
				//слипляем 2 строчки
				$objPHPExcel->getActiveSheet()->mergeCells($strC.':'.$strW);
				$i++;
			}
		}		
		
		//стили таблички
		$styleArray = array(
			'borders' => array(
				'outline' => array(
					'style' => PHPExcel_Style_Border::BORDER_THIN,
					'color' => array('argb' => 'FFFF0000'),
				),
			),
		);
		$stylerange = 'B23:AA'.$i;
		$objPHPExcel->getActiveSheet()->getStyle($stylerange)->applyFromArray($styleArray);
		//стили таблички ===============
		$inum = 22 + $i+4;
		$strB = 'B'.$inum;
		$objPHPExcel->getActiveSheet()->SetCellValue($strB, 'Всего наименований '.($i).', на сумму');		
		$inum = 22 + $i+5;
		$strB = 'B'.$inum;
		$objPHPExcel->getActiveSheet()->SetCellValue($strB, '');		
		
		// Rename sheet
		$objPHPExcel->getActiveSheet()->setTitle('invoice');
		
		// Save Excel 2007 file
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');		
		$objWriter->save("files/".$docname);
		unset($objPHPExcel);
		unset($objReader);
		if (is_file($docname)) unlink($docname);		
	}
	
}	
function change_docx($file, $docname, $leadData) {

	//	$dir = getcwd();
	//	echo $dir, "\n";
	//	echo $docname, "\n";
	//file_put_contents("file.txt", "1");
	
	// making a temp directory
	if (!is_dir("temp")) {
		mkdir("temp");
	} else {
		recursive_remove_directory("temp", true);
	}

    file_put_contents("file.txt", "2"); //add for test
	//safeify the directory
	//$edir = escapeshellarg($dir);

	//unzip everything
	//system("unzip $file -d temp");
    shell_exec("unzip $file -d temp");
	// replace the tokens
	$c = file_get_contents("temp/word/document.xml");

	$strres = updateValue($c, $leadData);
	unlink("temp/word/document.xml");
    unlink("file.txt");
    //@file_put_contents("temp/word/document1.xml",$c);
    @file_put_contents("temp/word/document.xml",$strres);
	// rezip
	if (is_file($docname)) unlink($docname);
    //выжные параметры для упаковки zip -> docx
    $toZip = array(
        "_rels",
        "docProps",
        "word",
        "[Content_Types].xml",
        "customXml",
    );

    $cmd = "cd temp && zip -r ../files/$docname ".implode(" ", $toZip);
	shell_exec($cmd);
    //$fdocxname = "/files/$docname";
    //rename($fdocxname,$fdocxname.'.docx');
    //system($cmd);

}
//2.======================================================================
if ($ftype == 'contract') {
    $newfilename2412 = 'dogovor.zip';
    //3.1=========================================================================
    $leadData = $_POST;
    //3.1.1 ======================================================================
    change_docx($template_file_name, $out_file_name, $leadData);
    //copy($template_file_name, "files/".$out_file_name);

    //$zip = new ZipArchive();
    //$zip->open($newfilename2412, ZipArchive::CREATE);
    //$zip->addFile("files/".$out_file_name.'.zip','Dogovor.docx'); //добавляет в архив dogovor.zip файл contract_template под именем Dogovor.docx
    //$zip->close();
    //$newfilename2412 = "dog-calgelunion.docx";
    $newfilename2412 = $out_file_name;
} elseif($ftype == 'invoice') {
    $leadData = $_POST;
    change_docx($template_file_name, $out_file_name, $leadData);
    $newfilename2412 = $out_file_name;
} elseif($ftype == 'act') {
    $leadData = $_POST;
    change_docx($template_file_name, $out_file_name, $leadData);
    $newfilename2412 = $out_file_name;	
} elseif($ftype == 'tradeinvoice') {	
	$leadData = $_POST;
    change_xlsx($template_file_name, $out_file_name, $leadData, $ftype);
    $newfilename2412 = $out_file_name;
} elseif($ftype =='tradeact') {
	$leadData = $_POST;
    change_xlsx($template_file_name, $out_file_name, $leadData, $ftype);
    $newfilename2412 = $out_file_name;
} 
else {
    //3.2=========================================================================
   

} //end else
//====4.Работа с гугл=========================================================
@file_put_contents("now-wc.txt","point4",FILE_APPEND);
require_once dirname(__FILE__).'/src/Google/autoload.php';
session_start();

$client_id = '493129122815-vur4ot74slh6l2piqqrpn34ara8m1hqs.apps.googleusercontent.com';
$client_secret = 'dDU3s5hfRrRYJQKsFLIBFEkU';
$redirect_uri = 'http://calgel.ru/esp/cb.php';

$client = new Google_Client();
//$client->setAuthConfigFile('client_secret.json');
$client->setClientId($client_id);
$client->setClientSecret($client_secret);
$client->setRedirectUri($redirect_uri);

//$client->setAuthConfigFile('client_secret.json');
$client->setAccessType("offline");
$client->addScope("https://www.googleapis.com/auth/drive");
$service = new Google_Service_Drive($client);

if (isset($_REQUEST['logout'])) {
    unset($_SESSION['upload_token']);
    @unlink("token");
}

if (isset($_GET['code'])) {
    $client->authenticate($_GET['code']);
    $_SESSION['upload_token'] = $client->getAccessToken();
    file_put_contents("token", $_SESSION['upload_token']);
    $redirect = 'http://' . $_SERVER['HTTP_HOST'] . $_SERVER['PHP_SELF'];
    header('Location: ' . filter_var($redirect, FILTER_SANITIZE_URL));
}

if (file_exists("token"))  {
    $tok = file_get_contents("token");
    $client->setAccessToken($tok);
    if ($client->isAccessTokenExpired()) {
        unset($_SESSION['upload_token']);
        @unlink("token");
    }
} else {
    $authUrl = $client->createAuthUrl();
}

if ($client->getAccessToken()) {
    @file_put_contents("now-wc.txt","point2.1 getAccessToken.file:"."files/".$newfilename2412,FILE_APPEND);
    $file = new Google_Service_Drive_DriveFile();

    //$file->setTitle($out_file_name);
    $file->setTitle($newfilename2412);

    if (($ftype == 'contract') or ($ftype == 'invoice') or ($ftype == 'act')) {
        @file_put_contents("now-wc.txt","point2.1.1 contract",FILE_APPEND);
        //$newfilename2412
        //array(
        //    'data' => file_get_contents("files/".$out_file_name),
        //    'mimeType' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        //    'uploadType' => 'media',
        //    'convert' => 'true'
        //)
        $result = $service->files->insert(
            $file,
            array(
                'data' => file_get_contents("files/".$newfilename2412),
                'mimeType' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'uploadType' => 'media',
                'convert' => 'true'
            )
        );

    } else {
        $result = $service->files->insert(
            $file,
            array(
                'data' => file_get_contents("files/".$newfilename2412),
                'mimeType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'uploadType' => 'media',
                'convert' => 'true'
            )
        );
    }

} elseif (isset($authUrl)) {
    @file_put_contents("now-wc.txt","point4.2",FILE_APPEND);
    $link = array(
        "auth_link"=>$authUrl,
    );
} else {
    @file_put_contents("now-wc.txt","point4.3",FILE_APPEND);
}


if (($ftype == 'contract') or ($ftype == 'invoice') or ($ftype == 'act')) {
	//docx
    if (!empty($result)){
        $link = array(
            "edit_link"=>$result->alternateLink,
            "dl_link"=>$result->exportLinks["application/vnd.openxmlformats-officedocument.wordprocessingml.document"]
        );
        //@unlink("files/".$out_file_name);
        @unlink("files/".$newfilename2412);
    }
} else {
	//xlsx
    if (!empty($result)){
        $link = array(
            "edit_link"=>$result->alternateLink,
            "dl_link"=>$result->exportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
        );
        @unlink("files/".$out_file_name);
    }

}
echo json_encode($link);
@file_put_contents("now-wc.txt","json:".json_encode($link),FILE_APPEND);
?>