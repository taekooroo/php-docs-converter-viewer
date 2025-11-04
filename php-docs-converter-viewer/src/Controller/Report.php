<?php
require_once __DIR__ . '/../vendor/autoload.php';
use Libs\ReportExportLibs;

$method = $_SERVER['REQUEST_METHOD'];

if (isset($_REQUEST['action'])) {
    switch ($_REQUEST['action'] ?? null) {
        case 'get_list':
            getReportList();
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
    $excel->setDataColumn($sheetData);
    $excel->streamExcelBlob($excel->getExcelBlob(), []);
}

/**
 * 엑셀 다운로드
 * @return void
 */
function downloadExcel() {

}

/**
 * PDF 다운로드
 * @return void
 */
function downloadPdf() {

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