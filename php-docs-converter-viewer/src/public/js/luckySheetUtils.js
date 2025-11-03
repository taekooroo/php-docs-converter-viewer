// https://dream-num.github.io/LuckysheetDocs/guide/

const defaultLuckysheetOptions = {
    
    // 숨길 행/열 범위 설정
    hiding: {
        // 사용자 옵션 지정하지 않으면 실행 안함
        startCol: 1,  // 숨기기 시작할 열 인덱스
        endCol: 0,   // 숨기기 종료할 열 인덱스
        startRow: 1, // 숨기기 시작할 행 인덱스
        endRow: 0    // 숨기기 종료할 행 인덱스
    },

    // Luckysheet 생성 시 사용될 UI 및 기능 설정
    options: {
        rowHeaderWidth: 0,          // 행 헤더 숨기기
        columnHeaderHeight: 0,      // 열 헤더 숨기기
        showinfobar: false,         // 상단 정보 바
        showsheetbar: false,        // 하단 시트 탭
        showstatisticBar: false,    // 하단 통계 바
        showZoom: false,
        allowEdit: false,           // 셀 편집 허용
        showstatisticBarConfig: {   // 통계 바 세부 설정
            zoom: true,             // 확대/축소 컨트롤
            count: false,           // 선택된 셀 개수 
	        view: false,            // 인쇄 미리보기
        },
        showtoolbar: false,         // 상단 도구 모음(툴바)
        enableAddRow: false,        // 행 추가 기능
        enableAddBackTop: false,    // '맨 위로 가기' 버튼
        sheetFormulaBar: false,     // 수식 입력창
        cellRightClickConfig: {     // 우클릭 메뉴
            copy: false,            // - 복사
            copyAs: false,          // - 다른 형식으로 복사
            paste: false,           // - 붙여넣기
            insertRow: false,       // - 행 삽입
            insertColumn: false,    // - 열 삽입
            deleteRow: false,       // - 행 삭제
            deleteColumn: false,    // - 열 삭제
            deleteCell: false,      // - 셀 삭제
            hideRow: false,         // - 행 숨기기
            hideColumn: false,      // - 열 숨기기
            rowHeight: false,       // - 행 높이 조절
            columnWidth: false,     // - 열 너비 조절
            clear: false,           // - 내용 지우기
            matrix: false,          // - 행렬 작업
            sort: false,            // - 정렬
            filter: false,          // - 필터
            chart: false,           // - 차트
            image: false,           // - 이미지
            link: false,            // - 하이퍼링크
            data: false,            // - 데이터 메뉴
            cellFormat: false       // - 셀 서식
        },
        columnlen : {
            0: 500,   // A열 너비 50px
            1: 60,   // B열 너비 60px
            5: 200   // F열 너비 200px
        }
    },

    // 붙여넣기 방지
    hook: {
        rangePasteBefore: function(range, data) {
            return false;
        },
    },

    zoomRatio: 1,       // 배율
    showGridLines: 0,   // 격자선 숨기기
};

/**
 * Luckysheet을 생성
 * @param {Blob} blobData - 변환할 Excel 파일의 Blob 데이터
 * @param {string} containerId - Luckysheet이 표시될 HTML 요소의 ID
 * @param {object} [userOptions={}] - 사용자가 재정의할 옵션 객체
 */
const displayExcelSheet = (blobData, containerId, userOptions = {}, imgDatas = {}) => {

    if (window.luckysheet) window.luckysheet.destroy();

    const mergedOptions = {
        hiding: { ...defaultLuckysheetOptions.hiding, ...(userOptions.hiding || {}) },
        options: {
            ...defaultLuckysheetOptions.options, 
            ...(userOptions.options || {}),
            cellRightClickConfig: {
                ...defaultLuckysheetOptions.options.cellRightClickConfig,
                ...(userOptions.options?.cellRightClickConfig || {})
            }
        },
        hook: { ...defaultLuckysheetOptions.hook, ...(userOptions.hook || {}) },
        zoomRatio: userOptions.zoomRatio ?? defaultLuckysheetOptions.zoomRatio,
        showGridLines: userOptions.showGridLines ?? defaultLuckysheetOptions.showGridLines,
    };

    LuckyExcel.transformExcelToLucky(blobData, (exportJson) => {
        if (!exportJson.sheets?.length) {
            alert('빈 파일이거나 변환에 실패했습니다.');
            return;
        }

        const sheet = exportJson.sheets[0];

        imgDataLength = Object.keys(imgDatas).length;
        if (imgDataLength > 0) {
            for (const key in imgDatas) {
                const imgData = imgDatas[key];
                // 시트에 삽입될 이미지 정보
                sheet.images = {
                    ...sheet.images,
                    ...setImageOptions(imgData.src, imgData.options)
                };
            }
        }

        // 숨길 열/행 설정
        const colhidden = {};
        const rowhidden = {};
        for (let i = mergedOptions.hiding.startCol; i <= mergedOptions.hiding.endCol; i++) colhidden[i] = 1;
        for (let i = mergedOptions.hiding.startRow; i <= mergedOptions.hiding.endRow; i++) rowhidden[i] = 1;

        const columnlen = {};
        const end = userOptions.columnEnd ?? null;
        const colWidth = userOptions.columnlen ?? null;
        if (colWidth !== null && end !== null) {
            for (let i = 0; i <= end; i++) {
                columnlen[i] = colWidth;
            }
        }

        sheet.config = { ...sheet.config, colhidden, rowhidden, columnlen };
        sheet.zoomRatio = mergedOptions.zoomRatio;
        sheet.showGridLines = mergedOptions.showGridLines;



        luckysheet.create({
            container: containerId,
            data: [sheet],
            ...mergedOptions.options,
            ...mergedOptions.hook,
        });
    });
};

/**
 * 
 * @param {*} imgData base64 이미지 데이터
 * @param {*} options 이미지 옵션
 * @returns 
 */
const setImageOptions = (imgData, options = {}) => {

    const uniqueId = `image_${Date.now()}_${Math.floor(Math.random() * 1000000)}`;
    const defaultOptions = { 
        "src": "", // 이미지 데이터 (base64 인코딩된 PNG)
        "fromCol": 0, // 이미지가 위치할 시작 열 인덱스
        "fromColOff": 0, // 이미지가 위치할 시작 열 오프셋 (px)
        "fromRow": 0, // 이미지가 위치할 시작 행 인덱스
        "fromRowOff": 0, // 이미지가 위치할 시작 행 오프셋 (px)
        "toCol": 2, // 이미지가 끝나는 열 인덱스
        "toColOff": 52, // 이미지가 끝나는 열 오프셋 (px)
        "toRow": 4, // 이미지가 끝나는 행 인덱스
        "toRowOff": 10, // 이미지가 끝나는 행 오프셋 (px)
        "originWidth": 194, // 원본 이미지 너비 (px)
        "originHeight": 110, // 원본 이미지 높이 (px)
        "type": "1", // 이미지 타입 (1: 일반 이미지)
        "isFixedPos": false, // 고정 위치 여부 (false: 셀 위치에 따라 이동)
        "fixedLeft": 0, // 이미지의 고정 좌측 위치 (px)
        "fixedTop": 0, // 이미지의 고정 상단 위치 (px)
        "border": { // 이미지 테두리 정보
            "color": "#000", // 테두리 색상
            "radius": 0, // 테두리 모서리 둥글기
            "style": "solid", // 테두리 스타일
            "width": 0 // 테두리 두께 (px)
        },
        "crop": { // 이미지 잘라내기 정보
            "height": 110, // 잘라낼 높이 (px)
            "offsetLeft": 0, // 좌측 오프셋 (px)
            "offsetTop": 0, // 상단 오프셋 (px)
            "width": 194 // 잘라낼 너비 (px)
        },
        "default": { // 기본 표시 정보
            "height": 90, // 표시 높이 (px)
            "left": 0, // 셀 내 좌측 위치 (px)
            "top": 0, // 셀 내 상단 위치 (px)
            "width": 150 // 표시 너비 (px)
        }
    };

    const mergedOptions = { ...defaultOptions, ...options };
    return { [uniqueId]: { ...mergedOptions, src: imgData } };
}