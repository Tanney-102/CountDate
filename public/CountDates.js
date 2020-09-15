const mainFile = document.getElementById('main-xlsx-file');
const subFile = document.getElementById('sub-xlsx-file');
const resultBtn = document.getElementById('result-btn');
const resultArea = document.getElementById('result-area');
const fileInputArea = document.getElementById('file-input-area');
const resultTableBody = document.getElementById("result-table-b");
const backBtn = document.getElementById('back-btn');
const fileName = document.getElementsByClassName('file-name');
const downloadBtn = document.getElementById('download-btn');
let mainSheet;
let subSheet;

mainFile.addEventListener('change', parseMainFile);
subFile.addEventListener('change', parseSubFile);
resultBtn.addEventListener('click', printResult);
backBtn.addEventListener('click', goBack);

function parseMainFile(e) {
    const reader = new FileReader();
    
    reader.onload = function() {
        const fileData =  reader.result;
        const wb = XLSX.read(fileData, {type : 'binary'});
        const firstSheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[firstSheetName];
        const mainHeader = getHeaderRow(sheet);
        const workerIdx = mainHeader.indexOf('아이돌보미명');
        const userIdx = mainHeader.indexOf('이용자명');
        const workTimeIdx = mainHeader.indexOf('활동시간');
        const workDateIdx = mainHeader.indexOf('이용시간');
        
        mainSheet = []; 
        let range = XLSX.utils.decode_range(sheet['!ref']);

        for(let R=2; R<=range.e.r; ++R) {
            let rowData = [];
            
            rowData.push(getElem(sheet, workerIdx, R));
            rowData.push(getElem(sheet, userIdx, R));
            rowData.push(getElem(sheet, workTimeIdx, R));
            rowData.push(getElem(sheet, workDateIdx, R));

            mainSheet.push(rowData);
        }
    }
    
    reader.readAsBinaryString(e.target.files[0]);
    fileName[0].innerHTML = e.target.files[0].name;
    console.log(mainSheet);
}

function parseSubFile() {
    const reader =  new FileReader();

    reader.onload = function() {
        const fileData = reader.result;
        const wb = XLSX.read(fileData, {type : 'binary'});
        const firstSheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[firstSheetName];
    
        subSheet = new Map(); // { worker => [user1, user2, ... ], }
        let range = XLSX.utils.decode_range(sheet['!ref']);

        for(let R = range.s.r+1; R<=range.e.r; ++R){
            let worker = sheet[XLSX.utils.encode_cell({c:range.s.c + 1, r:R})];
            let key = 'UNKNOWN';
            if(worker && worker.t) key = XLSX.utils.format_cell(worker);

            let value = [];
            for(let C = range.s.c + 2; C<=range.e.c; ++C){
                let user = sheet[XLSX.utils.encode_cell({c:C, r:R})];
                let tmpValue = 'UNKNOWN';
                if(user && user.t) tmpValue = XLSX.utils.format_cell(user);
                value.push(tmpValue);
            }
            
            subSheet.set(key, value);
        }

    }
    reader.readAsBinaryString(subFile.files[0]);
    fileName[1].innerHTML = subFile.files[0].name;
}

function getHeaderRow(sheet) {
    var headers = [];
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var C, R = range.s.r; /* start in the first row */
    /* walk every column in the range */
    for(C = range.s.c; C <= range.e.c; ++C) {
        var cell = sheet[XLSX.utils.encode_cell({c:C, r:R})] /* find the cell in the first row */

        var hdr = "UNKNOWN " + C; // <-- replace with your desired default 
        if(cell && cell.t) hdr = XLSX.utils.format_cell(cell);

        headers.push(hdr);
    }
    return headers;
}

function getElem(sheet, i, R){
    let tmp = sheet[XLSX.utils.encode_cell({c:i, r:R})]
    let elem = 'UNKNOWN';
    if(tmp && tmp.t) elem = XLSX.utils.format_cell(tmp);

    return elem;
}

function printResult() {
    if(mainSheet==undefined || subSheet==undefined)
        return;
    
    const result = getWorkDays();
    const workDays = result[0];
    const maskNum = result[1];
    
    // workDays.forEach( (v, i) => 
    for(let i=0; i<workDays.length; i++) {
        resultTableBody.innerHTML += '<tr><td class="t-body">' + (i + 1) + '</td>' +
                                     '<td class="t-body t-body-name">' + workDays[i][0] + '</td>' + 
                                     '<td class="t-body">' + workDays[i][1] + '</td>'+ 
                                     '<td class="t-body">' + maskNum[i][1] + '</td>'+ 
                                     '<td class="t-body">' + maskNum[i][1]*2000 + '</td></tr>';
    }

    fileInputArea.classList.toggle('hide');
    resultArea.classList.toggle('hide');
    downloadBtn.addEventListener('click', function() { download_XLSX(workDays, maskNum) });
}

function getWorkDays() {
    let workDays = new Map();
    let sortedWorkDays = [];
    let maskNum = new Map();
    let sortedMaskNum = [];

    mainSheet.forEach( v => {
        if(v[0] == 'UNKNOWN') return;
        if(!workDays.has(v[0]))
            workDays.set(v[0], 0);
        if(!maskNum.has(v[0]))
            maskNum.set(v[0], 0);
    });

    for(let date=1; date<32; date++) {
        let curRows = [];
        let repetition = new Set();

        // mainSheet
        // -> v[0]: 아이돌보미명, v[1]: 이용자명, v[2]:활동시간, v[3]: 이용일시
        mainSheet.forEach(v => {
            if(v[3].slice(8, 10)*1 == date)
                curRows.push(v);
        });

        // workDays
        curRows.forEach(v => {
            if(v[0]=='UNKNOWN' || v[1]=='UNKNOWN' || v[2] < 3 || repetition.has(v[0])) return;
            if(subSheet.has(v[0])){
                let exceptions = subSheet.get(v[0]);
                for(let i=0; i<exceptions.length; i++)
                    if(v[1] == exceptions[i]) return;
            }
            
            repetition.add(v[0]);
            workDays.set(v[0], workDays.get(v[0]) + 1);
        });

        repetition.clear();

        // maskNum
        curRows.forEach(v => {
            if(v[0]=='UNKNOWN' || v[1]=='UNKNOWN' || repetition.has(v[0])) return;
            repetition.add(v[0]);
            maskNum.set(v[0], maskNum.get(v[0]) + 1);
        });
    }

    for(let k of workDays.keys())
        sortedWorkDays.push([k, workDays.get(k)]);
    sortedWorkDays.sort();

    for(let k of maskNum.keys())
        sortedMaskNum.push([k, maskNum.get(k)]);
    sortedMaskNum.sort();

    return [sortedWorkDays, sortedMaskNum];
}

function goBack() {
    resultTableBody.innerHTML = null;

    fileInputArea.classList.toggle('hide');
    resultArea.classList.toggle('hide');
}

let excelHandler = {
    getExcelFileName : function() {
        return '교통비 정산.xlsx';
    },
    getSheetName : function() {
        return '교통비 정산';
    },
    getExcelData : function(workdays, maskNum) {
        let resultSheet = [['아이돌보미명', '교통비 지원일수(일)', '마스크 지원일수(일)','마스크 지원 금액(원)']];

        for(let i=0; i<workdays.length; i++) {
            let tmpArr = [];
            tmpArr.push(workdays[i][0]);
            tmpArr.push(workdays[i][1]);
            tmpArr.push(maskNum[i][1]);
            tmpArr.push(maskNum[i][1]*2000);

            resultSheet.push(tmpArr);
        }

        return resultSheet;
    },
    getWorksheet : function(workdays, maskNum) {
        return XLSX.utils.aoa_to_sheet(this.getExcelData(workdays, maskNum));
    }
}

function s2ab(s) { 
    var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
    var view = new Uint8Array(buf);  //create uint8array as viewer
    for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
    return buf;    
}

function download_XLSX(workdays, maskNum) {
    let wb = XLSX.utils.book_new();
    const newSheet = excelHandler.getWorksheet(workdays, maskNum);

    XLSX.utils.book_append_sheet(wb, newSheet, excelHandler.getSheetName());

    const wbout = XLSX.write(wb, {bookType:'xlsx', type: 'binary'});
    saveAs(new Blob([s2ab(wbout)], {type:'application/octet-stream'}), excelHandler.getExcelFileName());
}

