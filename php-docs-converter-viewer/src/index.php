<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>리포트 다운로드 시스템</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/css/pluginsCss.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/plugins.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/css/luckysheet.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/assets/iconfont/iconfont.css' />
    <link rel='stylesheet' href='public/css/style.css?v=<?php echo filemtime($_SERVER['DOCUMENT_ROOT'].'/public/css/style.css')?>' />
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
        <div class="row">
            <div class="col-lg-8 mb-4">
                <div class="card">
                    <div class="card-header bg-primary text-white">
                        <h5 class="card-title mb-0">
                            <i class="fas fa-eye me-2"></i>
                            데이터 미리보기
                        </h5>
                    </div>
                    <div class="card-body p-0">
                        <div id="preview-container" class="preview-container">
                            <div class="loading-overlay" id="loading-overlay">
                                <div class="text-center">
                                    <div class="spinner-border text-primary" role="status">
                                        <span class="visually-hidden">로딩 중...</span>
                                    </div>
                                    <p class="mt-2 text-muted">데이터를 불러오는 중...</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
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
                            <div class="mb-3">
                                <div class="bg-light p-3 rounded">
                                    <small class="text-muted">
                                        <strong>파일명:</strong> 리포트_<?php echo date('Y-m-d'); ?><br>
                                        <strong>생성일:</strong> <?php echo date('Y년 m월 d일 H:i'); ?><br>
                                        <strong>데이터 수:</strong> <span id="record-count"></span>개 레코드
                                    </small>
                                </div>
                            </div>
                            <div class="d-grid gap-2">
                                <button type="button" id="download-button" class="btn btn-primary btn-download" onclick="exportSettings()">
                                    <i class="fas fa-download me-2"></i>
                                    다운로드
                                </button>
                            </div>
                        </form>
                        <hr>
                        <div class="d-flex flex-column align-items-start gap-2">
                            <button type="button" class="btn btn-outline-secondary btn-sm" onclick="refreshPreview()">
                                <i class="fas fa-refresh me-2"></i>
                                미리보기 새로고침
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.7.1.js" integrity="sha256-eKhayi8LEQwp4NKxN+CfCh+3qOVUtJn3QNZ0TciWLP4=" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-mousewheel/3.1.13/jquery.mousewheel.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/luckysheet.umd.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/luckyexcel/dist/luckyexcel.umd.js"></script>
    <script src="public/js/luckySheetUtils.js?v=<?php echo filemtime($_SERVER['DOCUMENT_ROOT'].'/public/js/luckySheetUtils.js')?>"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
    
    <script>
        // 새로고침 함수
        const refreshPreview = () => {
            location.reload();
        };

        // Base64 문자열을 Blob으로 변환하는 함수
        const base64ToBlob = (base64String, mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') => {
            const byteCharacters = atob(base64String);
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {
                byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            const byteArray = new Uint8Array(byteNumbers);
            return new Blob([byteArray], { type: mimeType });
        };

        // 다운로드 함수
        const exportSettings = async () => {
            const selectedType = $('input[name="download_type"]:checked').val();
            try {
                const response = await fetch('/report', {
                    method: 'POST',
                    body: new URLSearchParams({ 
                        action: selectedType === "excel" ? "download_excel" : "download_pdf",
                    })
                });
                if (!response.ok) throw new Error("다운로드 실패");
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;

                // 파일명에 오늘 날짜 추가
                const today = new Date();
                const yyyy = today.getFullYear();
                const mm = String(today.getMonth() + 1).padStart(2, '0');
                const dd = String(today.getDate()).padStart(2, '0');
                const dateStr = `${yyyy}${mm}${dd}`;

                a.download = `${dateStr}.${selectedType === 'excel' ? 'xlsx' : 'pdf'}`;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
            } catch (err) {
                console.error('다운로드 중 오류 발생:', err);
                alert('다운로드 중 오류가 발생했습니다.');
            }
        };

        $(document).ready(() => {
            $.ajax({
                url: 'report',
                data: {'action': 'get_list'},
                method: 'GET',
                dataType: 'json',
                success: response => {
                    // 로딩 오버레이 숨기기
                    $('#loading-overlay').remove();
                    
                    // Base64 문자열을 Blob으로 변환
                    const excelBlob = base64ToBlob(response.excelBlob);
                    
                    const imgDatas = {};
                    if (response.imgDatas) {
                        imgDatas[0] = {src: response.imgDatas, options: {}};
                    }
                    const options = {
                        hiding: {
                            startCol: 10,
                            endCol: 59,
                            startRow: 42,
                            endRow: 84
                        },
                        columnEnd: 10, // 마지막 열
                        columnlen: 150, // width 조정
                        hook: {
                            // LuckySheet가 완전히 로드된 후 실행되는 콜백
                            workbookCreateAfter: () => {
                                // 모든 시트 데이터 가져오기
                                const sheets = luckysheet.getAllSheets();
                                if (sheets && sheets.length > 0) {
                                    const sheet = sheets[0];
                                    // celldata에서 실제 데이터가 있는 행의 개수를 계산
                                    if (sheet.celldata && sheet.celldata.length > 0) {
                                        // 행 번호를 기준으로 유니크한 행의 개수 계산 (헤더 제외)
                                        const uniqueRows = new Set();
                                        sheet.celldata.forEach(cell => {
                                            if (cell.r !== undefined) {
                                                uniqueRows.add(cell.r);
                                            }
                                        });
                                        const rowCount = Math.max(0, uniqueRows.size - 1);
                                        $('#record-count').text(rowCount);
                                    } else {
                                        $('#record-count').text('0');
                                    }
                                }
                            }
                        }
                    };
                    displayExcelSheet(excelBlob, 'preview-container', options, imgDatas);
                },
                error: (xhr, status, error) => {
                    console.error('초기화 중 오류 발생:', status, error);
                    console.error('응답:', xhr.responseText);
                    $('#loading-overlay').remove();
                    $('#preview-container').html('<div class="alert alert-danger m-3">데이터 로드 중 오류가 발생했습니다.</div>');
                }
            });
        });
    </script>
</body>
</html>