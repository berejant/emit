<?php

$config = array(
    'sourceExcel' => __DIR__ . '/xlsx/template.xlsx',
    'outputExcel' => __DIR__ . '/xlsx/%d.xlsx',
    'outputPdf'   => __DIR__ . '/pdf/%d.pdf',

    'cloudConvertApiKeys' => array(
        'XIU_E1PbC4IfV3g__voYH2cDEEAE3ESKUoRXOhhIN2N9uiwcyqyCTx9or-Tk5VxOAm3X3HU-9A3wA7aYfwYcPg',
        'FlPpDn4efRcAucXYUXFzk0a33O5WNfVHd1A-bB4_IUTUtIlJJP7tJML2_HTBGGkzRZQglK1YzfcImQoDPZYXUA',
        'qhr3LzW1LoMcxf1PL0QYLta8Isb3XfiEMSgyBod6_sldj01Abg6TeULnqemWeHfmr5coFubSZUOKVjSB-kHjsg'
    ),
);

ini_set('error_log', __DIR__ . '/error.log');
ini_set('log_errors', true);
ini_set('display_errors', false);
error_reporting(E_ALL);

$isAjaxRequest = isset($_SERVER['HTTP_X_REQUESTED_WITH']) && $_SERVER['HTTP_X_REQUESTED_WITH'] === 'XMLHttpRequest';

require __DIR__ . '/vendor/autoload.php';

$action = filter_input(INPUT_POST, 'action');

$answer = array(
    'status' => false
);

if ('getReport' === $action) {

    try {
        $variantNo = filter_input(INPUT_POST, 'variantNo', FILTER_VALIDATE_INT, array(
            'options' => array('min_range' => 21, 'max_range' => 45)
        ));

        if (!$variantNo) {
            throw new RequestException ('Не вказано номер варіанту');
        }

        $answer['link'] = generateReport($variantNo);
        $answer['status'] = true;

    } catch (RequestException $e) {
        $answer['error'] = $e->getMessage();
    } catch (Exception $e) {
        error_log($e->getMessage() . PHP_EOL . $e->getTraceAsString());

        $answer['error'] = 'Внутрішня помилка серверу';
    }

    if(!$isAjaxRequest) {
        header('Location: ' . $answer['link'], true, 301);
        exit;
    }
}

if ($isAjaxRequest) {
    header('Content-Type:application/json;charset=utf-8');
    echo json_encode($answer);
    exit;
}

/**
 * @param $variantNo integer
 */
function generateReport($variantNo)
{
    $inputExcelFilePath = genereteExcel($variantNo);

    $pdfFilePath = generatePdf($variantNo,  $inputExcelFilePath);

    if( __DIR__ === substr($pdfFilePath, 0, strlen(__DIR__)) ) {
        $pdfRelativeFilePath = substr($pdfFilePath, strlen(__DIR__) + 1);
    } else {
        $pdfRelativeFilePath = $pdfFilePath;
    }

    return $pdfRelativeFilePath;
}

/**
 * @param $variantNo integer
 */
function genereteExcel($variantNo)
{
    global $config;

    $outputExcelFilePath = sprintf($config['outputExcel'], $variantNo);

    $inputSourceExcel = $config['sourceExcel'];

    if (file_exists($outputExcelFilePath) && filemtime($outputExcelFilePath) > filemtime($inputSourceExcel)) {
        return $outputExcelFilePath;
    }

    $phpExcel = PHPExcel_IOFactory::load($inputSourceExcel, 'template.xlsx');

    $sheet = $phpExcel->getSheet(0);

    $sheet->setCellValue('C1', $variantNo);
    $sheet->setCellValue('C3', $variantNo - 19);

    unset($sheet);

    $objWriter = PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');

    $objWriter->save($outputExcelFilePath);

    return $outputExcelFilePath;
}

/**
 * @param $variantNo integer
 * @param $inputExcel string
 */
function generatePdf($variantNo, $inputExcelFilePath)
{
    global $config;

    $outputPdfFilePath = sprintf($config['outputPdf'], $variantNo);

    if (file_exists($outputPdfFilePath) && filemtime($outputPdfFilePath) && filemtime($inputExcelFilePath)) {
        return $outputPdfFilePath;
    }

    $api = new \CloudConvert\Api(getCloudConvertApiKey($variantNo));

    $api->convert(array(
        'inputformat' => 'xlsx',
        'outputformat' => 'pdf',
        'input' => 'upload',
        'filename' => 'input.xlsx',
        'file' => fopen($inputExcelFilePath, 'r'),
    ))
        ->wait()
        ->download($outputPdfFilePath);


    return $outputPdfFilePath;

}

/**
 * @param $variantNo index
 * @return string
 */
function getCloudConvertApiKey($variantNo)
{
    global $config;

    $count = count($config['cloudConvertApiKeys']);

    $index = $variantNo % $count;


    $key = $config['cloudConvertApiKeys'][$index];

    return $key;
}

class RequestException extends Exception
{

}
