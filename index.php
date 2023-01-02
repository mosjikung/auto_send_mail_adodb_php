<?php


require_once __DIR__ . '/vendor/phpmailer/phpmailer/src/Exception.php';
require_once __DIR__ . '/vendor/phpmailer/phpmailer/src/PHPMailer.php';
require_once __DIR__ . '/vendor/phpmailer/phpmailer/src/SMTP.php';

require 'vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

include 'adodb5/adodb.inc.php';

define("BASE_HATCH_HOST", "(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.74  )(PORT = 1521)))(CONNECT_DATA=(SID=NYTG)))");
define("BASE_HATCH_USER", "sf5");
define("BASE_HATCH_PASSWORD", "omsf5");



$conn = NewADOConnection('oci8');
$conn->setConnectionParameter('session_mode',OCI_SYSDBA);
$conn->setCharSet('utf8');
$conn->SetFetchMode(ADODB_FETCH_ASSOC);

$conn->connect(BASE_HATCH_HOST,BASE_HATCH_USER,BASE_HATCH_PASSWORD);

$_sql = "Select yarn_item,topdyed_color
            from QN_MAIL_YARN_LT_PUR";

$stmt = $conn->Prepare($_sql);
$res = $conn->Execute($stmt);
$row = $res->getAll();


if(!$row){
	exit();
}
echo "<pre>";
var_dump($row);
echo "</pre>";

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->fromArray(array_keys($row[0]),NULL,'A1');
$sheet->fromArray($row,NULL,'A2');

$writer = new Xlsx($spreadsheet);
 	$filename = 'Yarn_Lead_Time.xlsx';

if (is_writable($filename)) {
    unlink('Yarn_Lead_Time.xlsx');
} 
    $writer->save('Yarn_Lead_Time.xlsx');

	




$_sql2 = "select distinct send_mail_to from QN_MAIL_YARN_LT_PUR";
$stmt2 = $conn->Prepare($_sql2);
$res2 = $conn->Execute($stmt2);
$row2 = $res2->FetchRow();

echo "<pre>";
var_dump($row2["SEND_MAIL_TO"]);
echo "</pre>";

//$row_3 =  implode(",",$row2["SEND_MAIL_TO"]);
$data = explode(',',$row2["SEND_MAIL_TO"]);

echo "<pre>";
var_dump($data);
echo "</pre>";


    
//require_once __DIR__ . '/../setting.php';

/* use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception; */


/**
 * 
 */
 $_sql3 = "BEGIN SF5.DELETE_YARN_BOM_LT_LOG_MAIL; END;";
$stmt3 = $conn->Prepare($_sql3);
$res3 = $conn->Execute($stmt3); 

	$mail = new PHPMailer\PHPMailer\PHPMailer();

	// var_dump($mail);
	try {
		$mail->SMTPDebug = 0;
		$mail->isSMTP();
		$mail->SMTPOptions = array(
	        'ssl' => array(
	            'verify_peer' => false,
	            'verify_peer_name' => false,
	            'allow_self_signed' => true
	        )
	    );
		$mail->Host = 'smtp.nanyangtextile.com';
		$mail->SMTPAuth = false;
		$mail->Port = 25;
		$mail->SMTPKeepAlive = true;
        $mail->SMTPSecure = "TLS";
        $mail->CharSet = "UTF-8";
        $mail->Encoding = "base64";
        $mail->setFrom('admin_monitoring@nanyangtextile.com', 'Admin Monitoring');
        

        //$data2 = ['weerayooth.k@nanyangtextile.com','rattapon.s@nanyangtextile.com','aphiphu.s@nanyangtextile.com','wasin.m@nanyangtextile.com'];
         foreach ($data as $value) {
            $mail->addAddress($value);
        } 
  
      
        $mail->addAddress('weerayooth.k@nanyangtextile.com');
        $mail->addAttachment('Yarn_Lead_Time.xlsx');
        $mail->isHTML(true);
        $mail->Subject = "Auto Mail Warning Check Yarn LeadTime (Open QN)";

        $mail->Body = "Auto Mail Warning Check Yarn LeadTime (Open QN)";

        $mail->send();
		
		
		
		
		return true;
	} catch (Exception $e) {
		return false;
	}
	$conn->close();

?>