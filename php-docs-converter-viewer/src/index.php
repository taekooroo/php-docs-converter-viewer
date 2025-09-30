<?php
require_once 'vendor/autoload.php';
use Libs\ReportExportLibs;

// 다운로드 요청 처리
if (isset($_POST['download'])) {
    $downloadType = $_POST['download_type'] ?? 'excel';
    
    try {
        $reportExport = new ReportExportLibs($downloadType);
        
        // 샘플 데이터 생성 (실제 프로젝트에서는 데이터베이스에서 가져옴)
        $data = [
            'A1' => ['value' => '번호', 'options' => ['setBold' => true]],
            'B1' => ['value' => '이름', 'options' => ['setBold' => true]],
            'C1' => ['value' => '이메일', 'options' => ['setBold' => true]],
            'D1' => ['value' => '날짜', 'options' => ['setBold' => true]],
            'A2' => ['value' => '1', 'options' => ['type' => 'number']],
            'B2' => ['value' => '홍길동', 'options' => []],
            'C2' => ['value' => 'hong@example.com', 'options' => []],
            'D2' => ['value' => '2025-09-30', 'options' => []],
            'A3' => ['value' => '2', 'options' => ['type' => 'number']],
            'B3' => ['value' => '김철수', 'options' => []],
            'C3' => ['value' => 'kim@example.com', 'options' => []],
            'D3' => ['value' => '2025-09-30', 'options' => []],
            'A4' => ['value' => '3', 'options' => ['type' => 'number']],
            'B4' => ['value' => '이영희', 'options' => []],
            'C4' => ['value' => 'lee@example.com', 'options' => []],
            'D4' => ['value' => '2025-09-30', 'options' => []],
        ];
        
        // 데이터 설정
        $reportExport->setDataColumn($data);
        
        // Excel 또는 PDF 다운로드
        $reportExport->generate();
        
    } catch (Exception $e) {
        echo '<div class="alert alert-danger">다운로드 중 오류가 발생했습니다: ' . htmlspecialchars($e->getMessage()) . '</div>';
    }
}
?>

<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>리포트 다운로드 시스템</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        .preview-container {
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 0.375rem;
            min-height: 400px;
        }
        .download-section {
            background-color: white;
            border: 1px solid #dee2e6;
            border-radius: 0.375rem;
            padding: 20px;
        }
        .table-preview {
            font-size: 0.9em;
        }
        .btn-download {
            min-width: 120px;
        }
    </style>
</head>
<body class="bg-light">
    <div class="container mt-4">
        <!-- 헤더 -->
        <div class="row mb-4">
            <div class="col-12">
                <h1 class="h3 text-primary">
                    <i class="fas fa-file-download me-2"></i>
                    리포트 다운로드 시스템
                </h1>
                <p class="text-muted">데이터를 미리보고 Excel 또는 PDF 형식으로 다운로드하세요.</p>
            </div>
        </div>

        <!-- 메인 컨텐츠 -->
        <div class="row">
            <!-- 미리보기 컨테이너 -->
            <div class="col-lg-8 mb-4">
                <div class="card">
                    <div class="card-header bg-primary text-white">
                        <h5 class="card-title mb-0">
                            <i class="fas fa-eye me-2"></i>
                            데이터 미리보기
                        </h5>
                    </div>
                    <div class="card-body">
                        <div class="preview-container p-3">
                            <div class="table-responsive">
                                <table class="table table-striped table-hover table-preview">
                                    <thead class="table-dark">
                                        <tr>
                                            <th>번호</th>
                                            <th>이름</th>
                                            <th>이메일</th>
                                            <th>날짜</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td>1</td>
                                            <td>홍길동</td>
                                            <td>hong@example.com</td>
                                            <td>2025-09-30</td>
                                        </tr>
                                        <tr>
                                            <td>2</td>
                                            <td>김철수</td>
                                            <td>kim@example.com</td>
                                            <td>2025-09-30</td>
                                        </tr>
                                        <tr>
                                            <td>3</td>
                                            <td>이영희</td>
                                            <td>lee@example.com</td>
                                            <td>2025-09-30</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            <div class="text-center text-muted mt-3">
                                <small>총 3개의 레코드</small>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 다운로드 섹션 -->
            <div class="col-lg-4 mb-4">
                <div class="card">
                    <div class="card-header bg-success text-white">
                        <h5 class="card-title mb-0">
                            <i class="fas fa-download me-2"></i>
                            다운로드 옵션
                        </h5>
                    </div>
                    <div class="card-body">
                        <form method="POST" action="">
                            <input type="hidden" name="download" value="1">
                            
                            <!-- 파일 형식 선택 -->
                            <div class="mb-3">
                                <label class="form-label fw-bold">파일 형식 선택</label>
                                <div class="form-check">
                                    <input class="form-check-input" type="radio" name="download_type" id="excel" value="excel" checked>
                                    <label class="form-check-label" for="excel">
                                        <i class="fas fa-file-excel text-success me-2"></i>
                                        Excel (.xlsx)
                                    </label>
                                </div>
                                <div class="form-check">
                                    <input class="form-check-input" type="radio" name="download_type" id="pdf" value="pdf">
                                    <label class="form-check-label" for="pdf">
                                        <i class="fas fa-file-pdf text-danger me-2"></i>
                                        PDF (.pdf)
                                    </label>
                                </div>
                            </div>

                            <!-- 파일 정보 -->
                            <div class="mb-3">
                                <div class="bg-light p-3 rounded">
                                    <small class="text-muted">
                                        <strong>파일명:</strong> 리포트_<?php echo date('Y-m-d'); ?><br>
                                        <strong>생성일:</strong> <?php echo date('Y년 m월 d일 H:i'); ?><br>
                                        <strong>데이터 수:</strong> 3개 레코드
                                    </small>
                                </div>
                            </div>

                            <!-- 다운로드 버튼 -->
                            <div class="d-grid gap-2">
                                <button type="submit" class="btn btn-primary btn-download">
                                    <i class="fas fa-download me-2"></i>
                                    다운로드
                                </button>
                            </div>
                        </form>

                        <!-- 추가 기능 버튼들 -->
                        <hr>
                        <div class="d-grid gap-2">
                            <button type="button" class="btn btn-outline-secondary btn-sm" onclick="refreshPreview()">
                                <i class="fas fa-refresh me-2"></i>
                                미리보기 새로고침
                            </button>
                            <button type="button" class="btn btn-outline-info btn-sm" onclick="exportSettings()">
                                <i class="fas fa-cog me-2"></i>
                                내보내기 설정
                            </button>
                        </div>
                    </div>
                </div>

                <!-- 도움말 카드 -->
                <div class="card mt-3">
                    <div class="card-body">
                        <h6 class="card-title">
                            <i class="fas fa-info-circle text-info me-2"></i>
                            도움말
                        </h6>
                        <small class="text-muted">
                            • Excel: 데이터 편집 및 분석에 최적<br>
                            • PDF: 인쇄 및 공유에 최적<br>
                            • 파일은 자동으로 다운로드됩니다
                        </small>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
    
    <script>
        function refreshPreview() {
            // 미리보기 새로고침 기능
            location.reload();
        }

        function exportSettings() {
            // 내보내기 설정 모달 등을 여는 기능
            alert('내보내기 설정 기능을 구현하세요.');
        }

        // 폼 제출 시 로딩 표시
        document.querySelector('form').addEventListener('submit', function() {
            const button = this.querySelector('button[type="submit"]');
            button.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>처리 중...';
            button.disabled = true;
        });

        // 파일 형식 변경 시 정보 업데이트
        document.querySelectorAll('input[name="download_type"]').forEach(function(radio) {
            radio.addEventListener('change', function() {
                // 선택된 형식에 따른 정보 업데이트 로직
                console.log('선택된 형식:', this.value);
            });
        });
    </script>
</body>
</html>