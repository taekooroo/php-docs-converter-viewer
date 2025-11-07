<?php
require_once __DIR__ . '/../vendor/autoload.php';
use Libs\ReportExportLibs;

$method = $_SERVER['REQUEST_METHOD'];

if (isset($_REQUEST['action'])) {
    switch ($_REQUEST['action'] ?? null) {
        case 'get_list':
            getReportList();
            break;
        case 'download_excel':
            downloadExcel();
            break;
        case 'download_pdf':
            downloadPdf();
            break;
        default:
            echo "Unknown action.";
            break;
    }
} else {
    echo "No action specified.";
}

/**
 * 데이터 출력
 * @return void
 */
function getReportList() {
    $excel = new ReportExportLibs();
    $sheetData = dummyData();
    // 데이터 먼저 채우고 페이지 설정을 적용해야 전체 컬럼에 너비가 반영됩니다.
    $excel->setDataColumn($sheetData);
    $excel->configurePageSetup(pageSetting());
    $excel->streamExcelBlob($excel->getExcelBlob(), []);
}

/**
 * 엑셀 다운로드
 * @return void
 */
function downloadExcel() {
    $excel = new ReportExportLibs('excel');
    $sheetData = dummyData();
    // 데이터 먼저 채우고 페이지 설정을 적용해야 전체 컬럼에 너비가 반영됩니다.
    $excel->setDataColumn($sheetData);
    $excel->configurePageSetup(pageSetting());
    // $excel->drawImageLogo($imgData);
    $excel->setSheetTitle("보고서 샘플");
    $excel->generate();
}

/**
 * PDF 다운로드
 * @return void
 */
function downloadPdf() {
    $excel = new ReportExportLibs('pdf');
    $sheetData = dummyData();
    $excel->setDataColumn($sheetData);
    $excel->configurePageSetup(pageSetting('download_pdf'));
    // $excel->drawImageLogo($imgData);
    $excel->setSheetTitle("보고서 샘플");
    $excel->generate();

}

/**
 * 더미 데이터
 * @return array
 */
function dummyData() {
    $headers = ['번호', '이름', '이메일', '날짜'];
    $rows = [
        ['1', '홍길동', 'hong@example.com', '2025-09-30'],
        ['2', '김철수', 'kim@example.com', '2025-09-30'],
        ['3', '이영희', 'lee@example.com', '2025-09-30'],
        ['4', '테스트', 'test@example.com', '2025-09-30'],
    ];

    $data = [];
    $columns = range('A', 'D');

    foreach ($headers as $i => $header) {
        $cell = $columns[$i] . '1';
        $data[$cell] = [
            'value' => $header,
            'options' => ['setBold' => true]
        ];
    }

    foreach ($rows as $rowIdx => $row) {
        foreach ($row as $colIdx => $value) {
            $cell = $columns[$colIdx] . ($rowIdx + 2);
            $options = [];
            if ($colIdx === 0) {
                $options['type'] = 'number';
            }
            $data[$cell] = [
                'value' => $value,
                'options' => $options
            ];
        }
    }

    return $data;
}

/**
 * 엑셀 시트 설정
 * @return array
 */
function pageSetting($action = '') {
    $defaults = [
        'printArea' => 'A1:J40',
        'margins' => [
            'top' => 0.236,
            'bottom' => 0.236,
            'left' => 0.236,
            'right' => 0.236,
        ]
    ];

    $options = ($action === 'download_pdf')
        ? array_merge($defaults, [
            'colHeight' => 18,
            'colWidth' => 10,
            'wrapText' => true, // 자동 줄바꿈
            'margins' => array_merge($defaults['margins'], [
                'top' => 0.236,
            ])
        ])
        : array_merge($defaults, [
            'colHeight' => 18,
            'colWidth' => 25,
            'margins' => array_merge($defaults['margins'], [
                'top' => 0.787,
            ])
        ]);
    return $options;
}


function writeLog(string $message, string $logFilePath = __DIR__ . '/app.log'): void
{
    $timestamp = date('Y-m-d H:i:s');
    $logEntry = sprintf("[%s] %s\n", $timestamp, $message);
    file_put_contents($logFilePath, $logEntry, LOCK_EX);
}
