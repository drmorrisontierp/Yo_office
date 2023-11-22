/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */



Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    document.getElementById("input").onchange = () => tryCatch(updateLabel);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

const coinFlip = () => {
  const x = Math.round(Math.random()*100)
  if (x < 50) {
    return false
  } else {
    return true
  }
}

const randomProducts = (num) => {
  let products = [];
  while (products.length < 31) {
    let x = num < 6 ? 1 : 2;
    let n = (Math.round(Math.random() * (num - x)) + x).toString();
    let m = (Math.round(Math.random() * 8) + 2).toString();
    let product = `${n} \u00B7 ${m}`;
    let altProduct = `${m} \u00B7 ${n}`;
    if (!products.includes(product) && !products.includes(altProduct)) {
      coinFlip() ? products.push(product):products.push(altProduct);
    } else {
      continue;
    }
  }
  return products;
};

const populateTable = (num) => {
  let tableList = [];
  let products = randomProducts(num);
  for (let x = 0; x < 30; x += 2) {
    let rowList = [`${x + 1})`, products[x], " = ", " ", " ", `${x + 2})`, products[x + 1], " = ", " "];
    tableList.push(rowList);
  }
  return tableList;
};


async function insertHTML() {
  await Word.run(async (context) => {
    const num = document.getElementById("input").value;
    const tableData = populateTable(num);
    const widths = ["1.5cm", "2.8cm", "1.1cm", "2.7cm", "1cm", "1.5cm", "2.8cm", "1.1cm", "2.7cm"]
    let html =
      '<table style="font-size:20pt;">';
    for (let x = 0; x < tableData.length; x++) {
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
      rowHTML += "</tr>"
      html += rowHTML;
    }
    html += "</table>"
    console.log(html);
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
    blankParagraph.insertHtml(html, Word.InsertLocation.end);

    await context.sync();
  });
}

function updateLabel() {
  let value = document.getElementById("input").value;
  document.getElementById("input-label").innerHTML = value;
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