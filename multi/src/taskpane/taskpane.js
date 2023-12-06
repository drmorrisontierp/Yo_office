/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */



Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // bindings
document.getElementById("clear").onclick = () => tryCatch(clear);
document.getElementById("insert-facit").onclick = () => tryCatch(insertFacit);

document.getElementById("sub_one").onclick = () => tryCatch(insertSubOne);
document.getElementById("sub_two").onclick = () => tryCatch(insertSubTwo);
document.getElementById("sub_one_grid").onclick = () => tryCatch(insertSubOneGrid);
document.getElementById("sub_two_grid").onclick = () => tryCatch(insertSubTwoGrid);
document.getElementById("add_one").onclick = () => tryCatch(insertAddOne);
document.getElementById("add_two").onclick = () => tryCatch(insertAddTwo);
document.getElementById("add_one_grid").onclick = () => tryCatch(insertAddOneGrid);
document.getElementById("add_two_grid").onclick = () => tryCatch(insertAddTwoGrid);
document.getElementById("multi_one").onclick = () => tryCatch(insertMultiOne);
document.getElementById("multi_two").onclick = () => tryCatch(insertMultiTwo);
document.getElementById("multi_one_grid").onclick = () => tryCatch(insertMultiOneGrid);
document.getElementById("multi_two_grid").onclick = () => tryCatch(insertMultiTwoGrid);
document.getElementById("div_one").onclick = () => tryCatch(insertDivOne);
document.getElementById("div_one_no").onclick = () => tryCatch(insertDivOneNo);
document.getElementById("div_two").onclick = () => tryCatch(insertDivTwo);
document.getElementById("div_two_no").onclick = () => tryCatch(insertDivTwoNo);
document.getElementById("div_one_grid").onclick = () => tryCatch(insertDivOneGrid);
document.getElementById("div_two_grid").onclick = () => tryCatch(insertDivTwoGrid);
document.getElementById("multi").onclick = () => tryCatch(insertMultiplications);

    
    
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

//Code for slider
let num_one = 0;
let num_two = 70;

document.addEventListener("DOMContentLoaded", function() {
  let slider = document.getElementById("slider");
  let left = slider.querySelector("#left");
  let right = slider.querySelector("#right");
  left.style.left = "0px";
  right.style.left = "70px";

  const roundPointer = (pointer) => {
    let num = pointer.id == "left" ? num_one : num_two;
    let roundPointer = Math.round(num / 10) * 10;
    pointer.style.left = roundPointer.toString() + "px";
    if (pointer.id == "left") {
      num_one = roundPointer;
    } else {
      num_two = roundPointer;
    }
    //console.log(num_one, num_two, roundPointer)
  };

  left.addEventListener("mousedown", function(event) {
    event.preventDefault(); // prevent selection start (browser action)
    let shiftX = event.clientX - left.getBoundingClientRect().left;
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
    function onMouseMove(event) {
      let newLeft = event.clientX - shiftX - slider.getBoundingClientRect().left;
      if (newLeft < 0) {
        newLeft = 0;
      }
      if (newLeft > num_two - 40 && newLeft < 70) {
        //newLeft = num_two - 40;
        let oldVal = right.style.left;
        oldVal.replace(/[a-zA-Z]/g, "");
        let newVal = parseInt(oldVal) + 1;
        right.style.left = newVal.toString() + "px";
        num_two = newVal;
      }
      //let rightEdge = slider.offsetWidth - left.offsetWidth;
      //if (newLeft > rightEdge) {
      //  newLeft = rightEdge;
      //}
      if (newLeft > 70) {
        newLeft = 70;
      }
      left.style.left = newLeft + "px";
      num_one = newLeft;
    }
    function onMouseUp() {
      roundPointer(left);
      roundPointer(right);

      document.removeEventListener("mouseup", onMouseUp);
      document.removeEventListener("mousemove", onMouseMove);
    }
  });
  right.addEventListener("mousedown", function(event) {
    event.preventDefault(); // prevent selection start (browser action)
    let shiftX = event.clientX - right.getBoundingClientRect().left;
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
    function onMouseMove(event) {
      let newLeft = event.clientX - shiftX - slider.getBoundingClientRect().left;
      //if (newLeft < 0) {
      //  newLeft = 0;
      //}
      if (newLeft < 30) {
        newLeft = 30;
      }
      if (newLeft < num_one + 40 && newLeft > 40) {
        //newLeft = num_one + 40;
        let oldVal = left.style.left;
        oldVal.replace(/[a-zA-Z]/g, "");
        let newVal = parseInt(oldVal) - 1;
        left.style.left = newVal.toString() + "px";
        num_one = newVal;
      }
      let rightEdge = slider.offsetWidth - right.offsetWidth;
      if (newLeft > rightEdge - 5) {
        newLeft = rightEdge - 5;
      }
      right.style.left = newLeft + "px";
      num_two = newLeft;
    }
    function onMouseUp() {
      roundPointer(right);
      roundPointer(left);

      document.removeEventListener("mouseup", onMouseUp);
      document.removeEventListener("mousemove", onMouseMove);
    }
  });

  left.addEventListener("dragstart", function() {
    return false;
  });

  right.addEventListener("dragstart", function() {
    return false;
  });
});

//Global variables
let ordinal = 1;
let choice = "multi";
let facit = {};

const insertMultiplications = () => {
  choice = "multi";
  tryCatch(insertHTML);
};
const insertSubOne = () => {
  choice = "subOne";
  tryCatch(insertHTML);
};
const insertSubTwo = () => {
  choice = "subTwo";
  tryCatch(insertHTML);
};
const insertSubOneGrid = () => {
  choice = "subOneGrid";
  tryCatch(insertHTML);
};
const insertSubTwoGrid = () => {
  choice = "subTwoGrid";
  tryCatch(insertHTML);
};
const insertAddOne = () => {
  choice = "addOne";
  tryCatch(insertHTML);
};
const insertAddTwo = () => {
  choice = "addTwo";
  tryCatch(insertHTML);
};
const insertAddOneGrid = () => {
  choice = "addOneGrid";
  tryCatch(insertHTML);
};
const insertAddTwoGrid = () => {
  choice = "addTwoGrid";
  tryCatch(insertHTML);
};
const insertMultiOne = () => {
  choice = "multiOne";
  tryCatch(insertHTML);
};
const insertMultiTwo = () => {
  choice = "multiTwo";
  tryCatch(insertHTML);
};
const insertMultiOneGrid = () => {
  choice = "multiOneGrid";
  tryCatch(insertHTML);
};
const insertMultiTwoGrid = () => {
  choice = "multiTwoGrid";
  tryCatch(insertHTML);
};
const insertDivOne = () => {
  choice = "divOne";
  tryCatch(insertHTML);
};
const insertDivOneNo = () => {
  choice = "divOneNo";
  tryCatch(insertHTML);
};
const insertDivTwo = () => {
  choice = "divTwo";
  tryCatch(insertHTML);
};
const insertDivTwoNo = () => {
  choice = "divTwoNo";
  tryCatch(insertHTML);
};
const insertDivOneGrid = () => {
  choice = "divOneGrid";
  tryCatch(insertHTML);
};
const insertDivTwoGrid = () => {
  choice = "divTwoGrid";
  tryCatch(insertHTML);
};

const coinFlip = () => {
  const x = Math.round(Math.random() * 100);
  if (x < 50) {
    return false;
  } else {
    return true;
  }
};

const countList = (list, element) => {
  let count = 0;
  for (let item of list) {
    //console.log(element, item)
    if (element === item) count++;
  }
  return count;
};

const randomProducts = (numOne, numTwo) => {
  let products = [];
  while (products.length < 60) {
    let n = (Math.round(Math.random() * (numTwo - numOne)) + numOne).toString();
    let m = (Math.round(Math.random() * 8) + 2).toString();
    let product = `${n} \u00B7 ${m}`;
    let altProduct = `${m} \u00B7 ${n}`;
    if (!products.includes(product) && !products.includes(altProduct)) {
      products.push(ordinal);
      coinFlip() ? products.push(product) : products.push(altProduct);
      facit[ordinal] = parseInt(n) * parseInt(m);
      ordinal += 1;
    } else {
      continue;
    }
  }
  return products;
};

const populateTable = (numOne, numTwo) => {
  let tableList = [];
  let products = randomProducts(numOne, numTwo);
  for (let x = 0; x < 60; x += 2) {
    let rowList = [
      `${products[x]})`,
      products[x + 1],
      " = ",
      " ",
      " ",
      `${products[x + 2]})`,
      products[x + 3],
      " = ",
      " "
    ];
    tableList.push(rowList);
  }
  return tableList;
};

const populateSubAdd = (flag, num) => {
  let subOne = [[], [], []];
  let x = 0;
  while (x < num) {
    let t1 = Math.round(Math.random() * 8) + 1;
    let t2 = Math.round(Math.random() * 9);
    let t3 = Math.round(Math.random() * 9);
    let b1 = Math.round(Math.random() * 7) + 2;
    let b2 = Math.round(Math.random() * 9);
    let b3 = Math.round(Math.random() * 9);
    if (b1 > t1 && b2 == 0) {
      continue;
    }
    if (flag == 1) {
      if (b2 > t2) b2 = t2 >= 1 ? t2 - 1 : t2;
      if (b3 > t3) b3 = t3 >= 2 ? t3 - 2 : t3;
    }
    if (flag == 2) {
      if (b2 < t2 && b3 < t3) {
        if (coinFlip) {
          let memory = b2;
          b2 = t2;
          t2 = memory;
        } else {
          let memory = b3;
          b3 = t3;
          t3 = memory;
        }
      }
    }
    if (b1 >= t1) b1 = b1 - t1 > 0 ? t1 - 1 : 0;
    if (flag == 3) {
      if (b1 + t1 > 9) {
        b1 = Math.round(Math.random() * 4);
        t1 = Math.round(Math.random() * 4) + 1;
      }
      if (b2 + t2 > 9) {
        b2 = Math.round(Math.random() * 4);
        t2 = Math.round(Math.random() * 5);
      }
      if (b3 + t3 > 9) {
        b3 = Math.round(Math.random() * 4);
        t3 = Math.round(Math.random() * 5);
      }
    }
    let top = `${t1}${t2}${t3}`;
    let bottom = `${b1}${b2}${b3}`;
    if (parseInt(bottom) == 0) continue;

    if (top != bottom) {
      let topList = [];
      if (t1 == 0) {
        if (t2 == 0) {
          topList.push(t3);
        } else {
          topList.push(t2);
          topList.push(t3);
        }
      } else {
        topList.push(t1);
        topList.push(t2);
        topList.push(t3);
      }
      subOne[0].push(topList);
      let bottomList = [];
      if (b1 == 0) {
        if (b2 == 0) {
          bottomList.push(b3);
        } else {
          bottomList.push(b2);
          bottomList.push(b3);
        }
      } else {
        bottomList.push(b1);
        bottomList.push(b2);
        bottomList.push(b3);
      }

      subOne[1].push(bottomList);
      subOne[2].push(ordinal);
      if (flag < 3) {
        facit[ordinal] = parseInt(top) - parseInt(bottom);
      } else {
        facit[ordinal] = parseInt(top) + parseInt(bottom);
      }
      ordinal += 1;
    } else {
      continue;
    }
    x++;
  }
  return subOne;
};

const populateMulti = (flag) => {
  let x = 0;
  let last = flag == "short" ? 18 : 12;
  let multiList = [[], [], []];
  let checkList = [];
  while (x < last) {
    let t1 = Math.round(Math.random() * 1) + 1;
    let t2 = Math.round(Math.random() * 9);
    let t3 = Math.round(Math.random() * 9);
    let b1 = Math.round(Math.random() * 5) + 1;
    let b2 = Math.round(Math.random() * 7) + 2;
    let check = `${t1}${t2}${t3}`;
    if (checkList.includes(check)) continue;
    checkList.push(check);
    multiList[0].push([t1, t2, t3]);
    flag == "short" ? multiList[1].push([b2]) : multiList[1].push([b1, b2]);
    let product = parseInt(check) * parseInt(flag == "short" ? `${b2}` : `${b1}${b2}`);
    multiList[2].push(ordinal);
    facit[ordinal] = product;
    ordinal++;
    x++;
  }
  return multiList;
};

const populateDivision = (flag, amount) => {
  let quotients = [[], [], [], []];
  let check = [];
  while (quotients[2].length < amount) {
    let t1 = coinFlip() ? Math.round(Math.random() * 8) + 1 : 0;

    let t2 = t1 > 0 ? Math.round(Math.random() * 9) : Math.round(Math.random() * 8) + 1;
    let t3 = Math.round(Math.random() * 9);
    let b1 = Math.round(Math.random() * 7) + 2;
    if (countList(check, b1) > 2) {
      continue;
    }
    let num = t1 != 0 ? `${t1}${t2}${t3}` : `${t2}${t3}`;

    if (check.includes(num) && check.includes(b1)) continue;
    if (flag === "shortEasy") {
      if (parseInt(num) % b1 != 0) continue;
    }
    if (flag === "shortHard") {
      if ((parseInt(num) * 10000) % b1 != 0 || parseInt(num) % b1 == 0) continue;
    }

    check.push(num);
    check.push(b1);
    quotients[0].push(t1 ? [t1, t2, t3] : [t2, t3]);
    quotients[1].push([b1]);
    quotients[2].push(ordinal);
    let quotient = parseInt(num) / b1;
    facit[ordinal] = quotient;
    quotients[3].push(quotient.toString());
    ordinal++;
  }
  //console.log(quotients);
  return quotients;
};

const repeat = (word, amount) => {
  let sentence = "";
  let x = 0;
  while (x < amount) {
    sentence += word;
    x++;
  }
  return sentence;
};

const createMultiHtml = () => {
  let numOne = (num_one + 10) / 10;
  let numTwo = (num_two + 10) / 10;
  const tableData = populateTable(numOne, numTwo);
  const widths = ["1.5cm", "2.8cm", "1.1cm", "2.7cm", "1cm", "1.5cm", "2.8cm", "1.1cm", "2.7cm"];
  let html = `<div><p style="font-size:16pt;">Namn: ___________________ ${repeat(
    "&nbsp",
    18
  )} Tid: _________________</p></div><br/><table style="font-size:20pt;">`;
  for (let x = 0; x < tableData.length; x += 2) {
    let rowHTML = "<tr>";
    for (let y = 0; y < tableData[x].length; y++) {
      if (y == 3 || y == 8) {
        rowHTML += `<td style="width:${widths[y]}; height:1.2cm;text-align:center;border:solid black 1px">${tableData[x][y]}</td>`;
      } else if (y == 0 || y == 5) {
        rowHTML += `<td style="width:${widths[y]}; height:1.4cm;text-align:right;">${tableData[x][y]}</td>`;
      } else {
        rowHTML += `<td style="width:${widths[y]}; height:1.4cm;text-align:center;">${tableData[x][y]}</td>`;
      }
    }
    rowHTML += "</tr>";
    html += rowHTML;
  }
  html += "</table>";
  return html;
};

const algorithms = (flag, num) => {
  const style_clear = 'style="width:0.75cm;text-align:right;"';
  const style_clear_center = 'style="text-align:center;"';

  const style_text = 'style="width:0.75cm;text-align:right;"';
  const style_border_top = 'style="width:0.75cm;border-top:solid black 1px;text-align:right;"';
  const style_border_bottom = 'style="width:0.75cm;border-bottom:solid black 1px;text-align:right;"';
  const style_border_bottom_text = 'style="width:0.75cm;border-bottom:solid black 1px;text-align:right;"';
  const style_border_bottom_center_text = 'style="width:0.75cm;border-bottom:solid black 1px;text-align:center;"';
  const style_border_both =
    'style="width:0.75cm;border-bottom:solid black 1px;border-top:solid black 1px;text-align:right;"';

  const style_border_dots = 'style="width:0.75cm; height: 0.75cm;border:dotted #D3D3D3 1px;text-align:right;"';
  const style_grid = 'style="width:0.75cm; height: 0.75cm;border:dotted #D3D3D3 1px;text-align:center;"';
  let row = [0, 0];
  let table = 0;
  let spaces;
  let data;
  let sign, pattern, ui;
  let tableRows, tableColumns, gridColumns;
  let topOne, topTwo, topThree;
  let bottomOne, bottomTwo, bottomThree;
  let style, style0ne, styleTwo, styleThree;
  let helper;
  let topHtml, bottomHtml;

  switch (flag) {
    case "subNoBorrowing":
      data = populateSubAdd(1, num);
      pattern = "tba";
      style0ne = style_clear;
      styleTwo = style_text;
      styleThree = style_clear;
      tableRows = 6;
      tableColumns = 3;
      sign = "-";
      ui = "algorithm";
      break;
    case "subNoBorrowingGrid":
      data = populateSubAdd(1, num);
      pattern = "gggg";
      style = style_grid;
      gridColumns = 5;
      tableRows = 4;
      tableColumns = 3;
      sign = "-";
      ui = "grid";
      break;
    case "subBorrowing":
      data = populateSubAdd(2, num);
      pattern = "tba";
      style0ne = style_clear;
      styleTwo = style_text;
      styleThree = style_clear;
      tableRows = 6;
      tableColumns = 3;
      sign = "-";
      ui = "algorithm";
      break;
    case "subBorrowingGrid":
      data = populateSubAdd(2, num);
      pattern = "gggg";
      style = style_grid;
      gridColumns = 5;
      tableRows = 4;
      tableColumns = 3;
      sign = "-";
      ui = "grid";
      break;
    case "addEasy":
      data = populateSubAdd(3, num);
      pattern = "tba";
      style0ne = style_clear;
      styleTwo = style_text;
      styleThree = style_clear;
      tableRows = 6;
      tableColumns = 3;
      sign = "+";
      ui = "algorithm";
      break;
    case "addEasyGrid":
      data = populateSubAdd(3, num);
      gridColumns = 5;
      pattern = "gggg";
      style = style_grid;
      tableRows = 4;
      tableColumns = 3;
      sign = "+";
      ui = "grid";
      break;
    case "addHard":
      data = populateSubAdd(4, num);
      pattern = "tba";
      style0ne = style_clear;
      styleTwo = style_text;
      styleThree = style_clear;
      tableRows = 6;
      tableColumns = 3;
      sign = "+";
      ui = "algorithm";
      break;
    case "addHardGrid":
      data = populateSubAdd(4, num);
      pattern = "gggg";
      style = style_grid;
      gridColumns = 5;
      tableRows = 4;
      tableColumns = 3;
      sign = "+";
      ui = "grid";
      break;
    case "multiShort":
      data = populateMulti("short");
      pattern = "tba";
      style0ne = style_clear;
      styleTwo = style_text;
      styleThree = style_clear;
      tableRows = 6;
      tableColumns = 3;
      sign = "\u00B7";
      ui = "algorithm";
      break;
    case "multiShortGrid":
      data = populateMulti("short");
      pattern = "gggg";
      style = style_grid;
      gridColumns = 5;
      tableRows = 4;
      tableColumns = 3;
      sign = "\u00B7";
      ui = "grid";
      break;
    case "multiLong":
      data = populateMulti("long");
      pattern = "tbkka";
      style0ne = style_border_bottom;
      styleTwo = style_border_bottom_text;
      styleThree = style_border_both;
      tableRows = 4;
      tableColumns = 3;
      sign = "\u00B7";
      ui = "algorithm";
      break;
    case "multiLongGrid":
      data = populateMulti("long");
      pattern = "gggggg";
      style = style_grid;
      gridColumns = 5;
      tableRows = 3;
      tableColumns = 3;
      sign = "\u00B7";
      ui = "grid";
      break;
    case "divShortEasy":
      data = populateDivision("shortEasy", num);
      //console.log(data)
      helper = true;
      pattern = "d";
      gridColumns = 10;
      tableRows = 7;
      tableColumns = 2;
      sign = "/";
      ui = "algorithm";
      break;
    case "divShortEasyNo":
      data = populateDivision("shortEasy", num);
      //console.log(data)
      helper = false;
      pattern = "d";
      gridColumns = 10;
      tableRows = 7;
      tableColumns = 2;
      sign = "/";
      ui = "algorithm";
      break;
    case "divShortHard":
      data = populateDivision("shortHard", num);
      //console.log(data)
      helper = true;
      pattern = "d";
      gridColumns = 10;
      tableRows = 7;
      tableColumns = 2;
      sign = "/";
      ui = "algorithm";
      break;
    case "divShortHardNo":
      data = populateDivision("shortHard", num);
      //console.log(data)
      helper = false;
      pattern = "d";
      gridColumns = 10;
      tableRows = 7;
      tableColumns = 2;
      sign = "/";
      ui = "algorithm";
      break;
    case "divShortEasyGrid":
      data = populateDivision("shortEasy", num);
      pattern = "gg";
      style = style_grid;
      gridColumns = 10;
      tableRows = 6;
      tableColumns = 2;
      sign = "/";
      ui = "grid";
      break;
    case "divShortHardGrid":
      data = populateDivision("shortHard", num);
      pattern = "gg";
      style = style_grid;
      gridColumns = 10;
      tableRows = 6;
      tableColumns = 2;
      sign = "/";
      ui = "grid";
      break;
  }

  let html =
    '<div><div><p style="font-size:16;">Namn: __________________________</div><br/><table style="font-size: 16pt;">';

  for (let r = 0; r < tableRows; r++) {
    html += "<tr>";
    for (let c = 0; c < tableColumns; c++) {
      html += "<td>";
      if (ui == "grid") {
        let top = "";
        let bottom = "";
        for (let char of data[0][table]) {
          top += char.toString();
        }
        for (let char of data[1][table]) {
          bottom += char.toString();
        }
        html += `<div><p>${data[2][table]})${repeat("&nbsp", 3)} ${top} ${sign} ${bottom} </p></div>`;
      }
      html += `<table style="font-size:16pt;">`;
      for (let p = 0; p < pattern.length; p++) {
        if (pattern[p] == "g") {
          html += "<tr>";
          for (let ar = 0; ar < gridColumns; ar++) {
            html += `<td ${style}>&nbsp</td>`;
          }
          html += "</tr>";
        }
        if (pattern[p] == "d") {
          let answer = data[3][table];
          let answer_length = answer.length;
          let answer_grid = "";
          let gridUnit = `<td ${style_grid}>&nbsp</td>`;
          for (let x = 0; x < answer_length; x++) {
            if (answer[x] == ".") {
              answer_grid += `<td ${style_clear_center}>,</td>`;
            } else {
              answer_grid += gridUnit;
            }
          }

          if (data[0][table].length < 2) {
            topHtml = `
            <td ${style_border_bottom_text}>${data[0][table][0]}</td>
            `;
            bottomHtml = `<td ${style_clear_center}>${data[1][table][0]}</td>`;
          } else if (data[0][table].length < 3) {
            topHtml = `
            <td ${style_border_bottom_center_text}>${data[0][table][0]}</td>
            <td ${style_border_bottom_center_text}>${data[0][table][1]}</td>
            `;
            bottomHtml = `<td colspan="2" ${style_clear_center}>${data[1][table][0]}</td>`;
          } else {
            topHtml = `
            <td ${style_border_bottom_center_text}>${data[0][table][0]}</td>
            <td ${style_border_bottom_center_text}>${data[0][table][1]}</td>
            <td ${style_border_bottom_center_text}>${data[0][table][2]}</td>
            `;
            bottomHtml = `<td colspan="3" ${style_clear_center}>${data[1][table][0]}</td>`;
          }

          html += `
          <tr><td><table style="font-size: 16pt;">
          <td ${style_text}>${data[2][table]})</td>
          <td ${style_clear}>&nbsp</td>
          ${topHtml}
          </tr>
          <tr>
          <td ${style_clear}>&nbsp</td>
          <td ${style_clear}>&nbsp</td>
          ${bottomHtml}
          </tr></table></td>
          <td><table style="font-size: 16pt;><tr>
          <td ${style_clear}>=</td>
          <td ${style_clear}>&nbsp</td>
          `;
          html += helper ? answer_grid : repeat(gridUnit, 5);
          html += `</tr></table></td></tr>`;
        }

        if (pattern[p] == "t") {
          html += `
          <tr>
          <td ${style_text}>${data[2][table]})</td>
          <td ${style_clear}>&nbsp</td>
          <td ${style_text}>${data[0][table][0]}</td>
          <td ${style_text}>${data[0][table][1]}</td>
          <td ${style_text}>${data[0][table][2]}</td>
          </tr>`;
        }
        if (pattern[p] == "b") {
          if (data[1][table].length < 2) {
            bottomOne = "&nbsp";
            bottomTwo = "&nbsp";
            bottomThree = data[1][table][0];
          } else if (data[1][table].length < 3) {
            bottomOne = "&nbsp";
            bottomTwo = data[1][table][0];
            bottomThree = data[1][table][1];
          } else {
            bottomOne = data[1][table][0];
            bottomTwo = data[1][table][1];
            bottomThree = data[1][table][2];
          }

          html += `
          <tr>
          <td ${style0ne}>&nbsp</td>
          <td ${styleTwo}>${sign}</td>
          <td ${styleTwo}>${bottomOne}</td>
          <td ${styleTwo}>${bottomTwo}</td>
          <td ${styleTwo}>${bottomThree}</td>
          </tr>`;
        }
        if (pattern[p] == "a") {
          html += `
          <tr>
          <td ${styleThree}>&nbsp</td>
          <td ${style_border_both}>&nbsp</td>
          <td ${style_border_both}>&nbsp</td>
          <td ${style_border_both}>&nbsp</td>
          <td ${style_border_both}>&nbsp</td>
          </tr>
          `;
        }
        if (pattern[p] == "k") {
          html += `
          <tr>
          <td ${style_border_dots}>&nbsp</td>
          <td ${style_border_dots}>&nbsp</td>
          <td ${style_border_dots}>&nbsp</td>
          <td ${style_border_dots}>&nbsp</td>
          <td ${style_border_dots}>&nbsp</td>
          </tr>
          `;
        }
      }

      table++;
      html += '</table></td><td style="width:1cm;">&nbsp</td>';
    }
    if (r == tableRows - 1) {
      html += "</tr>";
    } else {
      html += '</tr><tr style="height:0.25cm; font-size:14pt"><td>&nbsp</td></tr>';
    }
  }
  html += "</table></div>";

  return html;
};

async function insertHTML() {
  await Word.run(async (context) => {
    const docBody = context.document.body;
    let html;
    if (choice == "multi") html = createMultiHtml();
    if (choice == "subOne") html = algorithms("subNoBorrowing", 18);
    if (choice == "subTwo") html = algorithms("subBorrowing", 18);
    if (choice == "addOne") html = algorithms("addEasy", 18);
    if (choice == "addTwo") html = algorithms("addHard", 18);

    if (choice == "subOneGrid") html = algorithms("subNoBorrowingGrid", 12);
    if (choice == "addOneGrid") html = algorithms("addEasyGrid", 12);
    if (choice == "subTwoGrid") html = algorithms("subBorrowingGrid", 12);
    if (choice == "addTwoGrid") html = algorithms("addHardGrid", 12);

    if (choice == "multiOne") html = algorithms("multiShort", 18);
    if (choice == "multiTwo") html = algorithms("multiLong", 12);
    if (choice == "multiOneGrid") html = algorithms("multiShortGrid", 12);
    if (choice == "multiTwoGrid") html = algorithms("multiLongGrid", 9);

    if (choice == "divOne") html = algorithms("divShortEasy", 14);
    if (choice == "divOneNo") html = algorithms("divShortEasyNo", 14);
    if (choice == "divTwo") html = algorithms("divShortHard", 14);
    if (choice == "divTwoNo") html = algorithms("divShortHardNo", 14);
    if (choice == "divOneGrid") html = algorithms("divShortEasyGrid", 12);
    if (choice == "divTwoGrid") html = algorithms("divShortHardGrid", 12);

    //console.log(html);
    await context.sync();
    docBody.insertHtml(html, Word.InsertLocation.end);
    docBody.insertBreak("Page", Word.InsertLocation.end);
    await context.sync();
    try {
      changeDocStyle();
    } catch (e) {
      console.log(e);
    }
    await context.sync();
  });
}

async function changeDocStyle() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("$none");

    await context.sync();
    paragraphs.items.forEach((item) => {
      if (item) {
        item.spaceAfter = 0;
        item.spaceBefore = 0;
      } else console.log("fail");
    });
    await context.sync();
    paragraphs.untrack();
    await context.sync();
  });
}

async function insertFacit() {
  await Word.run(async (context) => {
    const docBody = context.document.body;
    let text = "";
    let ordinalValues = Object.keys(facit);
    docBody.insertParagraph("Facit", Word.InsertLocation.end);
    //docBody.paragraphs.getLast().font.set({
    //  size: 14,
    //  underline: "Single"
    //});
    //docBody.paragraphs.getLast().spaceAfter = 20;
    await context.sync();
    for (let x = 1; x < ordinalValues.length + 1; x++) {
      if (x % 7 == 0 && x != 0) {
        text += `${ordinalValues[x - 1]}) ${facit[ordinalValues[x - 1]].toString().replace(".", ",")}`;
        docBody.insertParagraph(text, Word.InsertLocation.end);
        docBody.paragraphs.getLast().font.set({
          size: 14,
          underline: "None"
        });
        text = "";
        await context.sync();
      } else {
        text += `${ordinalValues[x - 1]}) ${facit[ordinalValues[x - 1]].toString().replace(".", ",")}\u0009`;
        docBody.paragraphs.getLast().font.set({
          size: 14,
          underline: "None"
        });
      }
    }
    docBody.insertParagraph(text, Word.InsertLocation.end);
    await context.sync();
    docBody.paragraphs.getLast().font.set({ size: 14 });
    docBody.paragraphs.getLast().spaceAfter = 20;
    await context.sync();
    docBody.insertBreak("Page", Word.InsertLocation.end);
    await context.sync();
  });
}

function updateLabel() {
  let value = document.getElementById("input").value;
  document.getElementById("input-label").innerHTML = value;
}

async function clear() {
  await Word.run(async (context) => {
    const docBody = context.document.body;
    facit = {};
    ordinal = 1;
    docBody.clear();
    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
