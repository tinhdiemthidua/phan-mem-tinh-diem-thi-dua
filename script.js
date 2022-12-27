//https://stackoverflow.com/questions/28892885/javascript-json-to-excel-file-download
let classGrade = [];
document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("header-grades").style.visibility = "hidden";

    let input = document.getElementById("stuList");
    let stuRows;
    // document.getElementById("stuName").addEventListener("change", ())
    input.addEventListener("change", function () {
        readXlsxFile(input.files[0]).then(function (rows) {
            stuRows = rows;
        });
    });
    document.getElementById("searchStuBtn").addEventListener("click", () => {
        let studentName = document.getElementById("stuName").value;

        searchStu(stuRows, studentName);
    });

    let week1 = document.getElementById("week1"),
        week2 = document.getElementById("week2"),
        week3 = document.getElementById("week3"),
        week4 = document.getElementById("week4");

    let dataWeek1, dataWeek2, dataWeek3, dataWeek4;
    let arrayClass = [];
    week1.addEventListener("change", function () {
        readXlsxFile(week1.files[0]).then(function (rows) {
            dataWeek1 = rows;
            dataWeek1.splice(0, 1);
            dataWeek1.forEach((e) => {
                arrayClass.push(e[0]);
                e.splice(0, 1);
            });
            console.log(dataWeek1);
        });
    });

    week2.addEventListener("change", function () {
        readXlsxFile(week2.files[0]).then(function (rows) {
            dataWeek2 = rows;
            dataWeek2.splice(0, 1);
            dataWeek2.forEach((e) => {
                e.splice(0, 1);
            });
        });
    });

    week3.addEventListener("change", function () {
        readXlsxFile(week3.files[0]).then(function (rows) {
            dataWeek3 = rows;
            dataWeek3.splice(0, 1);
            dataWeek3.forEach((e) => {
                e.splice(0, 1);
            });
        });
    });

    week4.addEventListener("change", function () {
        readXlsxFile(week4.files[0]).then(function (rows) {
            dataWeek4 = rows;
            dataWeek4.splice(0, 1);
            dataWeek4.forEach((e) => {
                e.splice(0, 1);
            });
        });
    });

    document.getElementById("month-report").addEventListener("click", () => {
        let dataMonth = monthlyReport(dataWeek1, dataWeek2, dataWeek3, dataWeek4);
        dataMonth.forEach((e) => {
            e.unshift(arrayClass.shift());
        });

        dataMonth.sort(sortFunction);
        dataMonth.unshift([
            "Tên lớp",
            "Điểm vệ sinh",
            "Điểm 15p",
            "Điểm thể dục",
            "Điểm nề nếp",
            "Điểm chuyên cần",
            "Điểm học tập",
            "Tổng điểm",
        ]);

        console.log(dataMonth);
        exportData("Báo cáo tháng.xlsx", dataMonth);
    });
});

function sortFunction(a, b) {
    return a[7] - b[7];
}

function addClass() {
    const week = document.getElementById("week").value;
    const duration = document.getElementById("duration").value;
    const className = document.getElementById("class-name").value;
    const vesinh = parseInt(document.getElementById("vesinh").value);
    const sh15 = parseInt(document.getElementById("sh15").value);
    const theduc = parseInt(document.getElementById("theduc").value);
    const nenep = parseInt(document.getElementById("nenep").value);
    const chuyencan = parseInt(document.getElementById("chuyencan").value);
    const hoctap = parseInt(document.getElementById("hoctap").value);
    const tong = parseInt(vesinh + sh15 + theduc + nenep + chuyencan + hoctap);

    document.getElementById("header-grades").style.visibility = "visible";

    if (
        week == "" ||
        duration == "" ||
        className == "" ||
        vesinh == "" ||
        sh15 == "" ||
        theduc == "" ||
        nenep == "" ||
        chuyencan == "" ||
        hoctap == ""
    ) {
        document.getElementById("error").textContent = "Chưa điền đủ thông tin";
        return;
    }

    document.getElementById("error").textContent = "";
    const classInfo = document.createElement("div");
    classInfo.innerHTML = `
    <div class="row" id="class-${className}">
        <div class="col">            
            <button class="btn btn-danger" data-class="${className}" id="delete">Xóa</button>
        </div>
        <div class="col">${className}</div>
        <div class="col">${vesinh}</div>
        <div class="col">${sh15}</div>
        <div class="col">${theduc}</div>
        <div class="col">${nenep}</div>
        <div class="col">${chuyencan}</div>
        <div class="col">${hoctap}</div>
        <div class="col">${tong}</div>
        <hr class="mt-2"/>
    </div>
    `;
    document.getElementById("show-list").appendChild(classInfo);

    classGrade.push({
        "Tên lớp": className,
        "Điểm vệ sinh": vesinh,
        "Điểm 15p": sh15,
        "Điểm thể dục": theduc,
        "Điểm nề nếp": nenep,
        "Điểm chuyên cần": chuyencan,
        "Điểm học tập": hoctap,
        "Tổng điểm": tong,
    });

    classGrade.sort(sortFunction);

    document.querySelectorAll("#delete").forEach((e) => {
        e.addEventListener("click", () => {
            classGrade = classGrade.filter((value) => value["Tên lớp"] != e.getAttribute("data-class"));
            document.getElementById(`class-${e.getAttribute("data-class")}`).style.display = "none";
        });
    });
    console.log(classGrade);
}

function exportData(filename, data) {
    var ws = XLSX.utils.json_to_sheet(data ? data : classGrade);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Báo cáo`);
    XLSX.writeFile(wb, filename);
}

function searchStu(rows, studentName) {
    rows.splice(0, 7);
    console.log(rows);
    rows.forEach((row) => {
        console.log(row);
        if (row[1] == studentName) {
            document.getElementById("studentInfo").innerHTML = `
        <div class="row">            
            <b class="col mb-2">Họ và tên</b>
            <b class="col mb-2">Mã học sinh</b>
            <b class="col mb-2">Ngày sinh</b>
            <b class="col mb-2">Giói tính</b>
            <b class="col mb-2">Dân tộc</b>
            <b class="col mb-2">Lớp năm học trước</b>
            <b class="col mb-2">Ghi chú</b>
        </div>
        <div class="row">            
            <div class="col">${row[1]}</div>
            <div class="col">${row[2]}</div>
            <div class="col">${row[3]}</div>
            <div class="col">${row[4]}</div>
            <div class="col">${row[5]}</div>
            <div class="col">${row[6] == null ? "" : row[6]}</div>                    
            <div class="col">${row[7] == null ? "" : row[7]}</div>                    
        </div>
            `;
        }
    });
}

function monthlyReport(dataWeek1, dataWeek2, dataWeek3, dataWeek4) {
    Array.prototype.matriceSum = function (a) {
        return this.reduce((p, c, i) => ((p[i] = c.reduce((f, s, j) => ((f[j] += s), f), p[i])), p), a.slice());
    };
    // console.log(dataWeek1.matriceSum(dataWeek2));
    let dataWeek12 = dataWeek1.matriceSum(dataWeek2);
    let dataWeek123 = dataWeek12.matriceSum(dataWeek3);
    let dataMonth = dataWeek123.matriceSum(dataWeek4);
    console.log(dataMonth);

    return dataMonth;
}

window.classGrade = classGrade;
