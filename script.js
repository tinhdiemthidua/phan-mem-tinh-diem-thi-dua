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
            "T??n l???p",
            "??i???m v??? sinh",
            "??i???m 15p",
            "??i???m th??? d???c",
            "??i???m n??? n???p",
            "??i???m chuy??n c???n",
            "??i???m h???c t???p",
            "T???ng ??i???m",
        ]);

        console.log(dataMonth);
        exportData("B??o c??o th??ng.xlsx", dataMonth);
    });
});

function sortFunction(a, b) {
    return a["T???ng ??i???m"] - b["T???ng ??i???m"];
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
        document.getElementById("error").textContent = "Ch??a ??i???n ????? th??ng tin";
        return;
    }

    document.getElementById("error").textContent = "";
    const classInfo = document.createElement("div");
    classInfo.innerHTML = `
    <div class="row" id="class-${className}">
        <div class="col">            
            <button class="btn btn-danger" data-class="${className}" id="delete">X??a</button>
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
        "T??n l???p": className,
        "??i???m v??? sinh": vesinh,
        "??i???m 15p": sh15,
        "??i???m th??? d???c": theduc,
        "??i???m n??? n???p": nenep,
        "??i???m chuy??n c???n": chuyencan,
        "??i???m h???c t???p": hoctap,
        "T???ng ??i???m": tong,
    });

    classGrade.sort(sortFunction);
    console.log(classGrade, classGrade.sort(sortFunction));

    document.querySelectorAll("#delete").forEach((e) => {
        e.addEventListener("click", () => {
            classGrade = classGrade.filter((value) => value["T??n l???p"] != e.getAttribute("data-class"));
            document.getElementById(`class-${e.getAttribute("data-class")}`).style.display = "none";
        });
    });
}

function exportData(filename, data) {
    var ws = XLSX.utils.json_to_sheet(data ? data : classGrade);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `B??o c??o`);
    XLSX.writeFile(wb, filename);
}

function searchStu(rows, studentName) {
    let found = false;
    rows.splice(0, 7);
    console.log(rows);
    rows.forEach((row) => {
        if (row[1] == studentName) {
            found = true;
            document.getElementById("studentInfo").innerHTML = `
        <div class="row">            
            <b class="col mb-2">H??? v?? t??n</b>
            <b class="col mb-2">M?? h???c sinh</b>
            <b class="col mb-2">Ng??y sinh</b>
            <b class="col mb-2">Gi??i t??nh</b>
            <b class="col mb-2">D??n t???c</b>
            <b class="col mb-2">L???p n??m h???c tr?????c</b>
            <b class="col mb-2">Ghi ch??</b>
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

    if (!found) document.getElementById("studentInfo").innerHTML = `<b>Kh??ng c?? h???c sinh n??o</b>`;
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
