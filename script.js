let excelData = [];
let studentPhotos = {};
let frontLogoSrc = "";
let backLogoSrc = "";
let frontBorderSrc = "";
let sign1Src = "";
let sign2Src = "";
let sign3Src = "";

// Subject Configurations per Class
const SUBJECT_MAP = {
    "PLAY": ["Hindi", "English", "Maths", "Urdu", "Deeniyat", "Poem", "P.T / Hygiene"],
    "LKG_UKG": ["Hindi", "English", "Maths", "Urdu", "Art", "Deeniyat", "Poem", "P.T / Hygiene"],
    "CLASS_1": ["Hindi", "English", "Maths", "Urdu", "Art", "General Knowledge", "Conversation", "Darse-a-Quran", "Deeniyat", "P.T / Hygiene"],
    "CLASS_2": ["Hindi", "English", "Maths", "Urdu", "Science", "Art", "General Knowledge", "Conversation", "Darse-a-Quran", "Deeniyat", "P.T / Hygiene"],
    "CLASS_3_5": ["Hindi", "English", "Maths", "Urdu", "Science", "S.S.T", "Computer", "Art", "General Knowledge", "Conversation", "Darse-a-Quran", "Deeniyat", "P.T / Hygiene"],
    "CLASS_6_8": ["Hindi", "English", "Maths", "Urdu", "Science", "S.S.T", "Computer", "Art", "General Knowledge", "Conversation", "Deeniyat", "P.T / Hygiene"]
};

// Template Downloader
function downloadTemplate() {
    const selectedClass = document.getElementById('classTemplateSelect').value;
    const subjects = SUBJECT_MAP[selectedClass];

    // Common Headers (No DOB)
    let headers = [
        "SrNo", "RollNo", "Name", "FatherName", "MotherName", "Class", "Address", "Attendance", "Remarks", "PromotedTo", "Rank"
    ];

    // Add Subjects
    subjects.forEach(sub => {
        headers.push(`${sub}_FA1`, `${sub}_SA1`, `${sub}_FA2`, `${sub}_SA2`);
    });

    const ws = XLSX.utils.aoa_to_sheet([headers]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `${selectedClass}_Result`);
    XLSX.writeFile(wb, `InfiTech_${selectedClass}_Format.xlsx`);
}

// Image Loaders
function handleImage(inputId, callback) {
    document.getElementById(inputId).addEventListener('change', function(e) {
        if(e.target.files[0]){
            const reader = new FileReader();
            reader.onload = function(ev) { callback(ev.target.result); };
            reader.readAsDataURL(e.target.files[0]);
        }
    });
}

handleImage('frontLogoInput', (res) => frontLogoSrc = res);
handleImage('backLogoInput', (res) => backLogoSrc = res);
handleImage('frontBorderInput', (res) => frontBorderSrc = res);
handleImage('sign1Input', (res) => sign1Src = res);
handleImage('sign2Input', (res) => sign2Src = res);
handleImage('sign3Input', (res) => sign3Src = res);

document.getElementById('photosInput').addEventListener('change', function(e) {
    const files = e.target.files;
    for (let i = 0; i < files.length; i++) {
        const fileName = files[i].name.split('.')[0];
        const reader = new FileReader();
        reader.onload = function(ev) { studentPhotos[fileName] = ev.target.result; };
        reader.readAsDataURL(files[i]);
    }
});

document.getElementById('excelInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(ev) {
        const data = new Uint8Array(ev.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        excelData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { raw: false });
    };
    reader.readAsArrayBuffer(file);
});

function generateReportCards() {
    if (excelData.length === 0) { alert("Data Empty. Please upload Excel."); return; }
    
    if(!backLogoSrc && frontLogoSrc) backLogoSrc = frontLogoSrc;
    if(!frontLogoSrc && backLogoSrc) frontLogoSrc = backLogoSrc;

    document.getElementById('ui-container').style.display = 'none';
    document.getElementById('preview-controls').style.display = 'flex';
    document.getElementById('output-area').style.display = 'block';

    const output = document.getElementById('output-area');
    const session = document.getElementById('sessionInput').value;
    output.innerHTML = ""; 

    excelData.forEach(student => {
        // Detect Subjects from Excel Column Names (Ending with _FA1)
        let subjects = [];
        Object.keys(student).forEach(key => {
            if (key.endsWith('_FA1')) {
                subjects.push(key.replace('_FA1', ''));
            }
        });

        // Ensure subjects follow the order in Excel if possible, or mapping
        // The loop above gets them in order of Excel columns usually.

        let rowsHTML = "";
        let subCount = 0;
        let sum_fa1=0, sum_sa1=0, sum_tot1=0;
        let sum_fa2=0, sum_sa2=0, sum_tot2=0;
        let sum_grand=0;

        subjects.forEach(sub => {
            let fa1 = parseFloat(student[`${sub}_FA1`]) || 0;
            let sa1 = parseFloat(student[`${sub}_SA1`]) || 0;
            let fa2 = parseFloat(student[`${sub}_FA2`]) || 0;
            let sa2 = parseFloat(student[`${sub}_SA2`]) || 0;
            
            let tot1 = fa1 + sa1;
            let tot2 = fa2 + sa2;
            let grand = tot1 + tot2;

            sum_fa1 += fa1; sum_sa1 += sa1; sum_tot1 += tot1;
            sum_fa2 += fa2; sum_sa2 += sa2; sum_tot2 += tot2;
            sum_grand += grand;
            subCount++;

            let d_fa1 = student[`${sub}_FA1`] || '-';
            let d_sa1 = student[`${sub}_SA1`] || '-';
            let d_fa2 = student[`${sub}_FA2`] || '-';
            let d_sa2 = student[`${sub}_SA2`] || '-';

            rowsHTML += `
            <tr>
                <td class="sub-name">${sub}</td>
                <td>${d_fa1}</td>
                <td>${d_sa1}</td>
                <td>${tot1 || '-'}</td>
                <td>${d_fa2}</td>
                <td>${d_sa2}</td>
                <td>${tot2 || '-'}</td>
                <td>${grand || '-'}</td>
            </tr>`;
        });

        let max_fa = subCount * 20;
        let max_sa = subCount * 80;
        let max_term = subCount * 100;
        let max_grand_total = subCount * 200;

        let percent = max_grand_total > 0 ? ((sum_grand / max_grand_total) * 100).toFixed(2) : 0;
        let resultStatus = percent > 33 ? "Passed" : "Failed";
        
        let photoSrc = studentPhotos[student.RollNo] || "https://via.placeholder.com/150?text=No+Photo";
        let frontImgTag = frontLogoSrc ? `<img src="${frontLogoSrc}" class="big-front-logo">` : '<h2>LOGO</h2>';
        let backImgTag = backLogoSrc ? `<img src="${backLogoSrc}" class="b-logo">` : `<div class="b-logo">LOGO</div>`;
        let wmTag = backLogoSrc ? `<img src="${backLogoSrc}" class="watermark">` : '';
        
        let borderHTML = frontBorderSrc ? `<img src="${frontBorderSrc}" class="front-border-bg">` : '';
        let designHTML = frontBorderSrc ? '' : `
            <div class="front-design-bg"></div>
            <div class="corner-design top-left"></div>
            <div class="corner-design bottom-right"></div>
        `;

        let s1 = sign1Src ? `<img src="${sign1Src}" class="sign-img">` : '';
        let s2 = sign2Src ? `<img src="${sign2Src}" class="sign-img">` : '';
        let s3 = sign3Src ? `<img src="${sign3Src}" class="sign-img pri-img">` : '';

        let backHeaderHTML = `
            <div class="back-header-row">
                ${backImgTag}
                <div class="b-center">
                    <h1 class="b-school-name">SITARA PUBLIC SCHOOL</h1>
                    <div class="b-address">Madina Colony Muzaffarnagar U.P.</div>
                    <div class="b-session">ACADEMIC REPORT (SESSION : ${session})</div>
                </div>
                ${backImgTag}
            </div>
        `;

        let reportCardHTML = `
        <div class="page">
            ${borderHTML}
            ${designHTML}
            
            <div class="front-content">
                <div class="school-header" style="margin-bottom: 30px;">
                    ${frontImgTag}
                </div>

                <div class="report-title">
                    Progress Report Card <br>
                    <span style="font-size: 20px; font-weight: normal;">Session (${session})</span>
                </div>
                
                <div class="student-photo-frame"><img src="${photoSrc}" alt="Student Photo"></div>
                
                <div class="profile-box">
                    <table class="profile-table">
                        <tr><td>Name</td><td>${student.Name}</td></tr>
                        <tr><td>Father's Name</td><td>${student.FatherName}</td></tr>
                        <tr><td>Mother's Name</td><td>${student.MotherName || '-'}</td></tr>
                        <tr><td>Address</td><td>${student.Address || '-'}</td></tr>
                        <tr><td>Class</td><td>${student.Class}</td></tr>
                        <tr><td>Roll No</td><td>${student.RollNo}</td></tr>
                        <tr><td>Sr No</td><td>${student.SrNo || '-'}</td></tr>
                    </table>
                </div>

                <div class="front-footer">
                    "Congratulation to all Students who secured good marks in their respective classes and those who aren't satisfied with their overall performance must give 100% the next time."
                </div>
            </div>
        </div>

        <div class="page">
            ${wmTag}
            ${backHeaderHTML}

            <table class="marks-header-table">
                <tr>
                    <td style="width:50%;">NAME: <span style="font-weight:normal; text-transform:uppercase;">${student.Name}</span></td>
                    <td style="width:50%;">FATHER: <span style="font-weight:normal; text-transform:uppercase;">${student.FatherName}</span></td>
                </tr>
                <tr>
                    <td>CLASS: <span style="font-weight:normal; text-transform:uppercase;">${student.Class}</span></td>
                    <td>ROLL NO: <span style="font-weight:normal;">${student.RollNo}</span></td>
                </tr>
            </table>

            <table class="marks-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="width: 25%;">SUBJECT</th>
                        <th>FA1<br><small>M.M. 20</small></th>
                        <th>SA1<br><small>M.M. 80</small></th>
                        <th>TOTAL<br><small>FA1 & SA1</small></th>
                        <th>FA2<br><small>M.M. 20</small></th>
                        <th>SA2<br><small>M.M. 80</small></th>
                        <th>TOTAL<br><small>FA2 & SA2</small></th>
                        <th>GRAND TOTAL<br><small>M.M. 200</small></th>
                    </tr>
                    <tr></tr>
                </thead>
                <tbody>
                    ${rowsHTML}
                    <tr class="max-marks-row">
                        <td style="text-align:right; padding-right:10px;">Total Max. Marks</td>
                        <td>${max_fa}</td>
                        <td>${max_sa}</td>
                        <td>${max_term}</td>
                        <td>${max_fa}</td>
                        <td>${max_sa}</td>
                        <td>${max_term}</td>
                        <td>${max_grand_total}</td>
                    </tr>
                    <tr class="obt-marks-row">
                        <td style="text-align:right; padding-right:10px;">Total Obtained</td>
                        <td>${sum_fa1}</td>
                        <td>${sum_sa1}</td>
                        <td>${sum_tot1}</td>
                        <td>${sum_fa2}</td>
                        <td>${sum_sa2}</td>
                        <td>${sum_tot2}</td>
                        <td>${sum_grand}</td>
                    </tr>
                </tbody>
            </table>

            <div class="footer-info">
                <div style="text-align:center; font-weight:bold; color:#d32f2f; margin-bottom:5px;">Evaluation Point</div>
                <div class="footer-grid">
                    <div style="border-right: 1px solid #ccc; padding-right: 15px;">
                        <div class="footer-row"><span>Remarks:</span> <b>${student.Remarks || 'Very Good'}</b></div>
                        <div class="footer-row"><span>Promoted to class:</span> <b>${student.PromotedTo || 'Next Class'}</b></div>
                        <div class="footer-row"><span>Result:</span> <b>${resultStatus}</b></div>
                        <div class="footer-row"><span>Percentage:</span> <b>${percent}%</b></div>
                    </div>
                    <div style="padding-left: 15px;">
                        <div class="footer-row"><span>Grace:</span> <b>-</b></div>
                        <div class="footer-row"><span>Attendance:</span> <b>${student.Attendance || '-'}</b></div>
                        <div class="footer-row"><span>Conduct:</span> <b>Good</b></div>
                        <div class="footer-row"><span>Rank:</span> <b>${student.Rank || '-'}</b></div>
                    </div>
                </div>
            </div>

            <div class="signatures">
                <div class="sign-box">
                    ${s1}
                    <div class="sign-line">Class Teacher Signature</div>
                </div>
                <div class="sign-box">
                    ${s2}
                    <div class="sign-line">Checked by</div>
                </div>
                <div class="sign-box">
                    ${s3}
                    <div class="sign-line">Principal<br>STAMP & SIGNATURE</div>
                </div>
            </div>

            <table class="grade-scale">
                <tr><td colspan="2" style="background:#e0e0e0; font-weight:bold;">GRADE DETAILS</td></tr>
                <tr><td>A+ (Outstanding) 90% and Above</td><td>A (Excellent) 80% to 89%</td></tr>
                <tr><td>B+ (Very Good) 70% to 79%</td><td>B (Good) 60% to 69%</td></tr>
                <tr><td>C+ (Average) 50% to 59%</td><td>C (Need improvement) 40% to 49%</td></tr>
                <tr><td colspan="2">D (Very Poor) 33% to 39% | E (Failed) 32% and Below</td></tr>
            </table>
        </div>`;
        output.innerHTML += reportCardHTML;
    });
}