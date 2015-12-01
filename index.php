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

    if (file_exists($outputExcelFilePath)) {
        return $outputExcelFilePath;
    }

    $phpExcel = PHPExcel_IOFactory::load($config['sourceExcel'], 'template.xlsx');

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
function generatePdf($variantNo, $inputExcel)
{
    global $config;

    $outputPdfFilePath = sprintf($config['outputPdf'], $variantNo);

    if (file_exists($outputPdfFilePath)) {
        return $outputPdfFilePath;
    }

    $api = new \CloudConvert\Api(getCloudConvertApiKey($variantNo));

    $api->convert(array(
        'inputformat' => 'xlsx',
        'outputformat' => 'pdf',
        'input' => 'upload',
        'filename' => 'input.xlsx',
        'file' => fopen($inputExcel, 'r'),
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

?>
<!doctype html>
<html lang="uk">
<head>
    <meta charset="utf-8">

    <title>Онлайн індивідуалка з управління витрататами</title>
    <meta name="description" content="Генератор PDF-звітів для індивідуального завдання з управління витратами в КНЕУ">
    <meta name="author" content="Anton Berezhnoj">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css"
          integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">

    <!--[if !IE]>-->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/zepto/1.1.6/zepto.min.js"></script>
    <!--<![endif]-->
    <!--[if IE]>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <![endif]-->

    <style>
        #getReportForm .row {
            text-align: center;
        }

        #reportGetLoading {
            display:none;
            margin-bottom: 20px;
        }

        #reportGetLoading p {
            font-size: 14px;
            line-height: 18px;
            font-family: monospace;
            vertical-align: baseline;
            margin: 0;
        }
    </style>

</head>

<body style="padding: 10px;">

<div class="container">
    <form class="jumbotron" id="getReportForm" method="post">
        <input type="hidden" name="action" value="getReport"/>

        <h1>Згенерувати звіт</h1>

        <p>для індивідуального завдання з управління витратами</p>

        <div id="reportGetError" style="display: none;" class="alert alert-danger" role="alert">...</div>

        <div class="row">
            <div class="form-group col-xs-3">
                <label for="varinatNo"">Номер варіанту:</label>
                <input type="number" class="form-control" id="varinatNo" name="variantNo" min="21" max="45"
                       placeholder="21...45">
            </div>
        </div>

        <div id="reportGetLoading" class="row">
            <div class="col-xs-3">
                <img src="loading.gif" width="42" height="42" />
                <p>Генерування...</p>
            </div>
        </div>


        <div class="row">
            <div class="form-group col-xs-3">
              <button type="submit" id="reportGetSubmitButton" class="btn btn-primary btn-lg">Згенерувати</button>
            </div>
        </div>
    </form>

    <hr>
    © 2015 Антон Бережний

</div>

</body>

<script>
    var $getReportForm = $('#getReportForm');
    var $reportGetError = $('#reportGetError');
    var $reportGetLoading = $('#reportGetLoading');
    var $reportGetSubmitButton = $('#reportGetSubmitButton');

    function reportGetAjaxStart () {
        $reportGetLoading.show();
        $reportGetError.hide();
        $reportGetSubmitButton.prop('disabled', true);
    }

    function reportGetAjaxComplete () {
        $reportGetLoading.hide();
        $reportGetSubmitButton.prop('disabled', false);
    }

    function getReportAjaxSuccess(answer) {
        reportGetAjaxComplete ();

        if (answer.error) {
            return showReportGetError(answer.error);
        }

        if (!answer.status) {
            return showReportGetError('Не вдалось виконати дію');
        }

        if(!answer.link) {
            return showReportGetError('Не знайдено посилання на файл');
        }

        location.assign(answer.link);
    }

    function getReportAjaxError(xhr, textError, errorThrown) {
        reportGetAjaxComplete ();

        showReportGetError('Не вдалось виконати запит до серверу (' + textError + ')');
    }

    function showReportGetError(error) {
        $reportGetError.text(error);
        $reportGetError.show();
    }


    $getReportForm.on('submit', function (event) {
        event.preventDefault();

        if($reportGetSubmitButton.is(':disabled')) {
            return;
        }

        reportGetAjaxStart();

        $.ajax({
            type: 'POST',
            url: location.href,
            data: $(this).serialize(),
            dataType: 'json',
            success: getReportAjaxSuccess,
            error: getReportAjaxError
        });

    });

</script>
</html>
