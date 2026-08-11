let groupedData;
let qrcodeText;

let qrcodeTextOptions = ["电器分装线", "内饰高架线", "预装线", "分装线", "越野车A线", "越野车B线", "越野车B-10线"];
if (localStorage.getItem("qrcodeTextOptions")) {
  qrcodeTextOptions = JSON.parse(localStorage.getItem("qrcodeTextOptions"));
  const options = document.getElementById("qrcodeTextOptions");
  qrcodeTextOptions.forEach((option) => {
    const opt = document.createElement("span");
    opt.innerHTML = option;
    opt.className = "qrcodeTextOption";
    opt.addEventListener("click", () => {
      document.getElementById("qrcodeText").value = option;
      qrcodeText = option;
    });
    options.appendChild(opt);
  });
} else {
  localStorage.setItem("qrcodeTextOptions", JSON.stringify(qrcodeTextOptions));
}

function arrGroup(arr, fn) {
  const obj = {};

  arr.forEach((item) => {
    const key = JSON.stringify(fn(item));

    obj[key] ??= [];

    obj[key].push(item);
  });

  return Object.values(obj).map((group) => group.sort((a, b) => b.totalAmount - a.totalAmount));
}

function createTable() {
  const quantityMode = document.getElementById("quantityMode").value;
  const content = document.getElementById("content");

  let page = null;

  let count = 0;

  function addLabel(label) {
    if (count % 6 === 0) {
      page = document.createElement("div");

      page.className = "page";

      content.appendChild(page);
    }

    page.appendChild(label);

    count++;
  }

  for (let i in groupedData) {
    let printCount = [];
    groupedData[i].forEach((material, index) => {
      printCount[index] = material.totalAmount;
    });

    for (let j = 0; j < groupedData[i][0].totalAmount; j++) {
      let table = document.createElement("table");
      let header = document.createElement("tr");
      const headers = ["图号", "数量", "计生号", "辆份"];
      headers.forEach((text) => {
        let th = document.createElement("th");

        th.textContent = text;

        header.appendChild(th);
      });
      table.appendChild(header);

      for (const index in groupedData[i]) {
        /*
                console.log("=================")
                console.log("index: " + index)
                console.log("currentCount: " + printCount[index])
                */

        if (printCount[index] <= 0) continue;

        const material = groupedData[i][index];
        let row = document.createElement("tr");

        let td1 = document.createElement("td");
        td1.textContent = material.materialNo;
        row.appendChild(td1);

        let td2 = document.createElement("td");
        td2.id = "quantity_" + i + "_" + j + "_" + index;

        if (quantityMode == "1") td2.textContent = 1;
        else if (quantityMode == "read") td2.textContent = material.totalAmount;
        else if (quantityMode == "dyn") {
          td2.textContent = material.totalAmount;
          td2.style.cursor = "pointer";
          td2.addEventListener("click", function () {
            this.textContent = this.textContent == 1 ? material.totalAmount : 1;
          });
        }

        row.appendChild(td2);

        const rowSpan = groupedData[i].filter((_, index) => printCount[index] > 0).length;
        if (index == 0) {
          let td3 = document.createElement("td");
          let span = document.createElement("span");
          span.style.fontSize = "x-large";
          span.style.fontWeight = "bold";
          span.textContent = material.planNo.split("-")[1];
          td3.appendChild(span);
          td3.appendChild(document.createElement("br"));
          td3.appendChild(document.createTextNode(material.planNo + "-" + String(material.no).padStart(4, "0")));
          td3.appendChild(document.createElement("br"));
          td3.appendChild(document.createTextNode(material.special));
          td3.rowSpan = rowSpan;
          row.appendChild(td3);
        }

        if (index == 0) {
          let td4 = document.createElement("td");
          td4.textContent = material.totalAmount + "-" + (j + 1);
          let img = document.createElement("img");
          // img.src = "https://picdl.sunbangyan.cn/2023/11/18/d7a30ef05dfe0b9736b015ccd8f340c4.jpg"
          img.src = "img.png";
          td4.appendChild(document.createElement("br"));
          td4.appendChild(img);
          td4.appendChild(document.createElement("br"));
          td4.appendChild(document.createTextNode(qrcodeText));
          td4.rowSpan = rowSpan + 1;
          row.appendChild(td4);
        }

        table.appendChild(row);
        printCount[index]--;
      }

      let last = document.createElement("tr");
      let td = document.createElement("td");
      td.colSpan = 3;
      td.textContent = groupedData[i][0].taskNo;
      last.appendChild(td);
      table.appendChild(last);
      let box = document.createElement("div");
      box.className = "aTable";
      box.appendChild(table);
      addLabel(box);
    }
  }
}

document.getElementById("fileInput").addEventListener("change", function (e) {
  let file = e.target.files[0];
  let reader = new FileReader();

  reader.onload = function (e) {
    let data = new Uint8Array(e.target.result);

    let workbook = XLSX.read(data, { type: "array" });

    let sheet = workbook.Sheets[workbook.SheetNames[0]];

    let rows = XLSX.utils.sheet_to_json(sheet);

    let jsonArray = rows.map((item) => ({
      taskNo: item["任务号"],

      materialNo: item["物料编码"],

      totalAmount: item["总量"],

      no: item["序号"],

      special: item["计生号"] || item["特殊描述"] || "",

      planNo: item["装配计划号"] || "",
    }));

    groupedData = arrGroup(jsonArray, (item) => item.taskNo + "&&" + item.special + "&&" + item.no + "&&" + item.planNo);
    console.log(groupedData);

    // qrcodeText = prompt("输入在二维码下方要显示的文字")
    qrcodeText = document.getElementById("qrcodeText").value;

    document.getElementById("menu").style.display = "none";

    createTable();

    if (!qrcodeTextOptions.includes(qrcodeText)) {
      qrcodeTextOptions.push(qrcodeText);
      localStorage.setItem("qrcodeTextOptions", JSON.stringify(qrcodeTextOptions));
    }
  };

  reader.readAsArrayBuffer(file);
});
