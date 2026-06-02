'use strict';

document.getElementById("nextBtn").addEventListener("click", function () {
    const selected = document.querySelector('input[name="listType"]:checked');

    if (!selected) {
        alert("一覧の種類を選んでください");
        return;
    }

    const type = selected.value;
    console.log("選ばれた一覧タイプ：", type);

    showFieldOptions(type);
});

// 一覧タイプに応じて項目チェックリストを表示
function showFieldOptions(type) {
    const area = document.getElementById("fieldSelectArea");

    let fields = [];

    if (type === "product") {
        fields = ["商品名", "商品画像", "サイズ", "入り数", "税抜き単価", "税抜き価格", "税込み価格"];
    } else if (type === "price") {
        fields = ["商品名", "税抜き単価", "税抜き価格", "税込み価格"];
    } else if (type === "catalog") {
        fields = ["商品名", "商品画像", "サイズ", "税込み価格"];
    } else if (type === "custom") {
        fields = ["商品名", "商品画像", "サイズ", "入り数", "税抜き単価", "税抜き価格", "税込み価格", "備考"];
    }

    let html = "<h3>必要な項目を選んでください</h3>";

    fields.forEach(f => {
        html += `<label><input type="checkbox" class="fieldCheck" value="${f}"> ${f}</label><br>`;
    });

    area.innerHTML = html;

    // チェックが変わるたびにフォームを再生成
    document.querySelectorAll(".fieldCheck").forEach(chk => {
        chk.addEventListener("change", generateForm);
    });
}

function generateForm() {
    const checks = document.querySelectorAll(".fieldCheck:checked");
    const formArea = document.getElementById("formArea");

    if (checks.length === 0) {
        formArea.innerHTML = "";
        return;
    }

    const selectedFields = Array.from(checks).map(c => c.value);

    let html = `
        <h3>入力行</h3>
        <div id="rowsArea"></div>
        <button id="addRowBtn">＋ 行を追加</button>
        <button id="downloadBtn">Excelを生成</button>
    `;

    formArea.innerHTML = html;

    // 最初の1行を追加
    document.getElementById("rowsArea").appendChild(createRow(selectedFields));

    // 行追加ボタン
    document.getElementById("addRowBtn").addEventListener("click", () => {
        document.getElementById("rowsArea").appendChild(createRow(selectedFields));
    });

    // Excel生成ボタン
    document.getElementById("downloadBtn").addEventListener("click", generateExcel);
}

function generateExcel() {
    const rows = document.querySelectorAll(".rowItem");

    if (rows.length === 0) {
        alert("入力行がありません");
        return;
    }

    const excelData = [];
    let headers = [];

    rows.forEach((row, index) => {
        const inputs = row.querySelectorAll(".inputField");
        const rowData = {};

        inputs.forEach(input => {
            const key = input.dataset.name;
            const value = input.value;

            rowData[key] = value;

            if (index === 0) {
                headers.push(key);
            }
        });

        excelData.push(Object.values(rowData));
    });

    excelData.unshift(headers);

    // Excel生成
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(excelData);

    // ★ メインカラー取得
    const mainColor = document.querySelector('input[name="mainColor"]:checked')?.value || "#004A77";
    const colorHex = mainColor.replace("#", "").toUpperCase();

    // ★ ヘッダーに色をつける
    headers.forEach((h, i) => {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: i });
        if (!ws[cellAddress]) return;

        ws[cellAddress].s = {
            fill: { fgColor: { rgb: colorHex } },
            font: { color: { rgb: "FFFFFF" }, bold: true },
            alignment: { horizontal: "center" }
        };
    });

    // シートにスタイルを適用
    wb.Sheets["Sheet1"] = ws;

    XLSX.writeFile(wb, "generated_list.xlsx");
}


function createRow(selectedFields) {
    const row = document.createElement("div");
    row.className = "rowItem";

    let html = "";

    selectedFields.forEach(field => {

        // 画像
        if (field === "商品画像") {
            html += `<input type="file" data-name="${field}" class="inputField" placeholder="${field}">`;
        }

        // 単価・入り数（自動計算対象）
        else if (field === "入り数" || field === "税抜き単価") {
            html += `<input type="number" data-name="${field}" class="inputField calcTarget" placeholder="${field}">`;
        }

        // 税抜き価格 → 入力できる（calcTarget に含める）
        else if (field === "税抜き価格") {
            html += `<input type="number" data-name="${field}" class="inputField calcTarget" placeholder="${field}">`;
        }

        // 税込み価格 → 自動計算専用
        else if (field === "税込み価格") {
            html += `<input type="number" data-name="${field}" class="inputField autoCalc" placeholder="${field}" disabled>`;
        }

        // その他
        else {
            html += `<input type="text" data-name="${field}" class="inputField" placeholder="${field}">`;
        }
    });

    html += `<button class="deleteRowBtn">－</button>`;
    row.innerHTML = html;

    // 削除ボタン
    row.querySelector(".deleteRowBtn").addEventListener("click", () => {
        row.remove();
    });

    // 自動計算イベント
    row.querySelectorAll(".calcTarget").forEach(input => {
        input.addEventListener("input", () => autoCalcRow(row));
    });

    return row;
}

function autoCalcRow(row) {
    const unit = row.querySelector('[data-name="税抜き単価"]');
    const qty = row.querySelector('[data-name="入り数"]');
    const ex = row.querySelector('[data-name="税抜き価格"]');
    const inc = row.querySelector('[data-name="税込み価格"]');

    // 税抜き単価が無くても、税抜き価格 → 税込み価格 の計算は行う
    const u = unit ? Number(unit.value) : null;
    const q = qty ? Number(qty.value) || 1 : 1;

    // 単価 × 入り数 → 税抜き価格
    if (unit && !isNaN(u) && !isNaN(q)) {
        const exPrice = u * q;
        if (ex) ex.value = exPrice;
    }

    // 税抜き価格 → 税込み価格（単価が無くても動く）
    if (ex) {
        const exPrice = Number(ex.value);
        if (!isNaN(exPrice)) {
            const incPrice = Math.round(exPrice * 1.1);
            if (inc) inc.value = incPrice;
        }
    }
}






