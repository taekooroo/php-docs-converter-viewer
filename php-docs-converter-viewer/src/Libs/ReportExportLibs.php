<?php
namespace Libs;

use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ReportExportLibs {

    private Spreadsheet $spreadsheet;
    private Worksheet $sheet;
    private ?string $downloadType;
    private $PDF_OUTPUT_DIR;
    private $configurePageSetuped = false;
    private $db;
    private $insertIdx;
    private $fileSize;

    /**
     * 생성자
     * @param string|null $downloadType 'excel', 'pdf', BLOB = null
     */
    public function __construct(?string $downloadType = null, $db = null) {
        
        $this->PDF_OUTPUT_DIR = dirname(__DIR__) . "/view/excel_templates/"; // pdf 임시 저장 경로
        $this->downloadType = $downloadType;
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->db = $db;
    }

    public function __destruct() {
        if ($this->db != null) {
            $this->db->close();
            $this->db = null;
        }
    }

    /**
     * 엑셀 데이터 및 다운로드 호출
     */
    public function generate(): void {
        // 타입별 엑셀 다운로드
        if ($this->downloadType === 'excel') {
            $this->downloadExcel();
        } elseif ($this->downloadType === 'pdf') {
            $this->downloadPdf();
        }
    }

    /**
     * 엑셀 파일 바이너리 데이터(BLOB)를 생성 및 반환
     * @return string 엑셀 바이너리 데이터
     */
    public function getExcelBlob(): string {
        $writer = new Xlsx($this->spreadsheet);
        
        ob_start();
        $writer->save('php://output');
        $excelBlob = ob_get_contents();
        ob_end_clean();

        return $excelBlob;
    }

    /**
     * 엑셀 파일 바이너리 데이터(BLOB)를 JSON 형태로 스트림 출력
     * @param string $excelBlob 엑셀 바이너리 데이터
     */
    public function streamExcelBlob(string $excelBlob, array $imgData = []): void {
        // HTTP 헤더 설정
        if (ob_get_length()) {
            ob_end_clean();
        }
        header('Content-Type: application/json; charset=utf-8');
        header('Cache-Control: no-cache, must-revalidate');
        header('Pragma: no-cache');

        if (empty(!$imgData)) {
            $extension = $imgData['extension'];
            $imageBase64 = "data:image/{$extension};base64," . base64_encode(file_get_contents($_SERVER['DOCUMENT_ROOT'].$imgData['path']));
        }

        $arrResult = [
            'excelBlob' => base64_encode($excelBlob),
            'imgDatas' => $imageBase64,
        ];

        echo json_encode($arrResult, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    }
    
    /**
     * 데이터 세팅
     */
    public function setDataColumn(array $arrData): void {
        foreach ($arrData as $cell => $data) {
            $this->setFormattedCell($cell, $data['value'] ?? '', $data['options']);
        }

        if (!$this->configurePageSetuped) {
            $this->configurePageSetup();
        }
    }
    
    /**
     * 셀 서식 지정
     * @param string $cellAddress
     * @param mixed $value
     * @param string $formatType
     */
    private function setFormattedCell(string $cellAddress, $value, array $options = []): void {
        $formatType = strtolower($options['type'] ?? 'string');
        $fontSize = $options['fontsize'] ?? 9;

        // 기본값 설정
        $dataType = DataType::TYPE_STRING;
        $formatCode = NumberFormat::FORMAT_TEXT;
        $horizontal = Alignment::HORIZONTAL_CENTER;

        switch ($formatType) {
            case 'number':
                $dataType = DataType::TYPE_NUMERIC;
                $formatCode = NumberFormat::FORMAT_NUMBER;
                break;
            case 'number_00':
                $dataType = DataType::TYPE_NUMERIC;
                $formatCode = NumberFormat::FORMAT_NUMBER_00;
                break;
            case 'comma':
                $dataType = DataType::TYPE_NUMERIC;
                $formatCode = '#,##0';
                $horizontal = Alignment::HORIZONTAL_RIGHT;
                break;
            default:
                break;
        }

        // 정렬 옵션 처리
        if (isset($options['align'])) {
            $alignOption = strtolower($options['align']);
            if ($alignOption === 'left') {
                $horizontal = Alignment::HORIZONTAL_LEFT;
            } elseif ($alignOption === 'right') {
                $horizontal = Alignment::HORIZONTAL_RIGHT;
            } else {
                $horizontal = Alignment::HORIZONTAL_CENTER;
            }
        }

        

        $this->sheet->setCellValueExplicit($cellAddress, $value, $dataType); // 셀에 값 설정
        $this->sheet->getStyle($cellAddress)->getNumberFormat()->setFormatCode($formatCode); // 셀의 숫자 형식 설정
        $this->sheet->getStyle($cellAddress)->getAlignment()->setHorizontal($horizontal); // 가로 정렬
        $this->sheet->getStyle($cellAddress)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER); // 세로 중앙 정렬

        $this->sheet->getStyle($cellAddress)->getFont()->setSize($fontSize); // 폰트 사이즈 설정
        if (!empty($options['setBold']) && $options['setBold'] === true) {
            $this->sheet->getStyle($cellAddress)->getFont()->setBold(true);
        } else {
            $this->sheet->getStyle($cellAddress)->getFont()->setBold(false);
        }

        // 테두리 설정
        if (!empty($options['borderBottom']) && $options['borderBottom'] === true) {
            $this->sheet->getStyle($cellAddress)->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);
        }
        if (!empty($options['borderLeft']) && $options['borderLeft'] === true) {
            $this->sheet->getStyle($cellAddress)->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THIN);
        }
        if (!empty($options['borderRight']) && $options['borderRight'] === true) {
            $this->sheet->getStyle($cellAddress)->getBorders()->getRight()->setBorderStyle(Border::BORDER_THIN);
        }
    }

    public function setRangeBorder(string $range, array $options = []): void {
        $borderStyle = $options['style'] ?? Border::BORDER_THIN;
        $borderColorArr = ['argb' => 'FF000000'];

        $colorObj = new \PhpOffice\PhpSpreadsheet\Style\Color($borderColorArr['argb'] ?? 'FF000000');
        $this->sheet->getStyle($range)->getBorders()->getOutline()
            ->setBorderStyle($borderStyle)
            ->setColor($colorObj);
    }

    public function setCellOptions(string $cellRange, string $type, array $options = []): void {
        switch ($type) {
            case 'borderTopThin':
                // 모든 테두리 얇게
                $this->sheet->getStyle($cellRange)->getBorders()->getTop()->setBorderStyle(Border::BORDER_THIN);
                break;
            case 'borderTopMedium':
                // 상단 두꺼운 테두리
                $this->sheet->getStyle($cellRange)->getBorders()->getTop()->setBorderStyle(Border::BORDER_MEDIUM);
                break;
            case 'borderOutlineThin':
                // 모든 테두리 얇게
                $this->sheet->getStyle($cellRange)->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
                break;
            case 'backgroundColor':
                $fillColor = $options['fillColor'] ?? 'FFFF00'; // 기본 노란색
                
                $this->sheet->getStyle($cellRange)->getFill()->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setARGB($fillColor);
                break;
            default:
                break;
        }

         // 테두리 설정
        if ($type === 'borderTopMedium') {
            $this->sheet->getStyle($cellRange)->getBorders()->getTop()->setBorderStyle(Border::BORDER_MEDIUM);
        }
    }

    /**
     * PDF 변환 인쇄 및 페이지 설정
     */
    public function configurePageSetup(array $options = []): void {
        $printArea = $options['printArea'] ?? 'A1:J40';
        $colHeight = $options['colHeight'] ?? 18;
        $colWidth = $options['colWidth'] ?? 9;
        $margins = $options['margins'] ?? [
            'top' => 0.6,
            'bottom' => 0.6,
            'left' => 0.6,
            'right' => 0.6,
        ];

        // 가로방향 가운데 정렬
        $this->sheet->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
        // 세로방향 상단 정렬
        $this->sheet->getPageSetup()->setHorizontalCentered(true);
        //눈금선 숨기기
        $this->sheet->setShowGridlines(false);

        $this->sheet
            ->getPageSetup()
            ->setPrintArea($printArea)
            ->setOrientation(PageSetup::ORIENTATION_PORTRAIT) // 세로 방향
            ->setPaperSize(PageSetup::PAPERSIZE_A4) // A4 용지
            ->setFitToWidth(1) // 가로에 맞춤
            ->setFitToHeight(0) // 세로 무제한
            ;
        
        // 셀 너비 및 행 높이 설정
        for ($row = 1; $row <= $this->sheet->getHighestRow(); $row++) {
            $this->sheet->getRowDimension($row)->setRowHeight($colHeight);
        }
        $highestColumn = Coordinate::columnIndexFromString($this->sheet->getHighestColumn());
        for ($col = 1; $col <= $highestColumn; $col++) {
            $columnLetter = Coordinate::stringFromColumnIndex($col);
            $this->sheet->getColumnDimension($columnLetter)->setWidth($colWidth);
        }

        // 여백 설정
        $this->sheet->getPageMargins()
            ->setTop($margins['top'])->setBottom($margins['bottom'])
            ->setLeft($margins['left'])->setRight($margins['right']);


        $this->configurePageSetuped = true;
    }
    
    /**
     * 엑셀 다운로드
     */
    private function downloadExcel(): void {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="인보이스_엑셀.xlsx"');
        header('Cache-Control: max-age=0');

        $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;
    }

    public function drawImageLogo(array $imgData = []): void {
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setDescription('Company Logo');
        $drawing->setPath($_SERVER['DOCUMENT_ROOT'].$imgData['path']); // 이미지 파일 경로
        $drawing->setResizeProportional(true);
        $drawing->setWidth($drawing->getWidth() * 0.7);
        $drawing->setHeight($drawing->getHeight() * 0.7);
        $drawing->setCoordinates('A1'); // 삽입할 셀 위치
        $drawing->setWorksheet($this->sheet); // 워크시트 지정
    }

    /**
     * Excel to PDF Dounload
     */
    private function downloadPdf(): void {
        $uniqueId = uniqid('export_', true);
        $xlsxFile = $this->PDF_OUTPUT_DIR . "{$uniqueId}.xlsx";
        $pdfOutputDir = $this->PDF_OUTPUT_DIR . date('Ym');
        $pdfFile = $pdfOutputDir . "/{$uniqueId}.pdf";

        if (!is_dir($this->PDF_OUTPUT_DIR)) {
            mkdir($this->PDF_OUTPUT_DIR, 0755, true);
        }

        // 엑셀 파일 저장
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($xlsxFile);

        try {
            $this->convertToPdf($xlsxFile, $pdfOutputDir);

            if (file_exists($pdfFile)) {
                header('Content-Type: application/pdf');
                header('Content-Disposition: attachment; filename="INVOICE.pdf"');
                header('Content-Length: ' . filesize($pdfFile));
                $result = readfile($pdfFile);
            } else {
                throw new Exception("PDF 파일 생성에 실패했습니다.");
            }
        } catch (Exception $e) {
            http_response_code(500);
            // 실제 운영 환경에서는 로그를 남기는 것이 좋습니다.
        } finally {
            // 임시 파일 삭제
            if (file_exists($xlsxFile)) {
                @unlink($xlsxFile);
            }
            if (file_exists($pdfFile)) {
                @unlink($pdfFile);
            }
            exit;
        }
    }
    
    /**
     * LibreOffice를 사용하여 XLSX 파일을 PDF로 변환
     * @param string $sourceXlsxPath
     * @param string $outputDir
     * @throws Exception
     */
    private function convertToPdf(string $sourceXlsxPath, string $outputDir): void {
        // PDF 출력 디렉토리 생성
        if (!is_dir($outputDir)) {
            mkdir($outputDir, 0755, true);
        }

        $cmd = sprintf(
            '/usr/bin/libreoffice --headless --convert-to pdf --outdir %s %s',
            escapeshellarg($outputDir),
            escapeshellarg($sourceXlsxPath)
        );

        $descriptorSpec = [
            1 => ['pipe', 'w'], // stdout
            2 => ['pipe', 'w'], // stderr
        ];
        
        $env = $_ENV;
        $env['HOME'] = '/tmp'; // LibreOffice 실행을 위한 HOME 환경 변수 설정

        $process = proc_open($cmd, $descriptorSpec, $pipes, null, $env);

        if (is_resource($process)) {
            $stdout = stream_get_contents($pipes[1]);
            fclose($pipes[1]);
            $stderr = stream_get_contents($pipes[2]);
            fclose($pipes[2]);
            $return_var = proc_close($process);

            if ($return_var !== 0) {
                throw new Exception("PDF 변환 실패 (LibreOffice 오류)\nStderr: $stderr\nStdout: $stdout");
            }
        } else {
            throw new Exception("PDF 변환 프로세스를 실행할 수 없습니다.");
        }
    }
}