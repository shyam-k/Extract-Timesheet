// Copyright (c) 2011 The Chromium Authors. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

// Called when the user clicks on the browser action.
chrome.browserAction.onClicked.addListener(function(tab) {
  // No tabs or host permissions needed!
  console.log('Extracting timesheet from  ' + tab.url + ' : tab !!!!');
  chrome.tabs.executeScript({
    code:codetext
      });
});

/*  chrome.webNavigation.onCompleted.addListener(function() {
      alert("This is my favorite website!");
  }, {url: [{urlMatches : 'http://ims.pitsolutions.com'}]});
*/


var convertToCsv = function convertToCsv(){




	if (document.getElementById("PreviousEntrytasks") === null) {
	    alert("No Timesheet Info Found On This Page");
	}

  var table = document.getElementById("PreviousEntrytasks");
  for (var i = 0; i < table.rows.length; i++) {
   // table.rows[i].cells[5].remove(); 
    //table.rows[i].cells[5].remove(); 
    //table.rows[i].cells[5].remove(); 
    var colIndex1 = 4;
    var colIndex2 = 3;
    var row = table.rows[i];
    var cell1 = row.cells[colIndex1];
    var cell2 = row.cells[colIndex2];
    var siblingCell1 = row.cells[Number(colIndex1) + 1];
    //row.insertBefore(cell1, cell2);
    //row.insertBefore(cell2, siblingCell1);



  }


	var count = document.getElementById("PreviousEntrytasks").rows.length;
	var content = "";
	for (i = 0; i < count; ++i) {
	    content = "\"" + 
       document.getElementById("PreviousEntrytasks").rows[i].cells[0].innerHTML + 
       "\",\"" + document.getElementById("PreviousEntrytasks").rows[i].cells[4].innerHTML + 
       "\",\"" + document.getElementById("PreviousEntrytasks").rows[i].cells[1].innerHTML + 
       "\",\"" + document.getElementById("PreviousEntrytasks").rows[i].cells[3].childNodes[0].innerHTML + "\" \r " + content;
	}
    newWindow = window.open("", null, "height=50,width=200,status=yes,toolbar=no,menubar=no,location=no");
    const MIME_TYPE = "text/plain";
    var bb = new Blob([content], {
        type: MIME_TYPE
    });
    var a = document.createElement("a");
    a.download = "timesheet.csv";
    a.href = window.URL.createObjectURL(bb);
    a.textContent = "Download Timesheet";
    a.dataset.downloadurl = [MIME_TYPE, a.download, a.href].join(":");
    newWindow.document.body.appendChild(a);
    newWindow.document.title = "Timesheet";	
}

var codetext = '';
codetext = codetext + convertToCsv + 'convertToCsv();';



function deleteColumns (table, colIndex1) {
  if (table && table.rows ) {
    for (var i = 0; i < table.rows.length; i++) {
      table.rows[i].cells[colIndex1].remove(); 
    }
  }
}

function swapColumns (table, colIndex1, colIndex2) {
  if (table && table.rows && table.insertBefore && colIndex1 != colIndex2) {
    for (var i = 0; i < table.rows.length; i++) {
      var row = table.rows[i];
      var cell1 = row.cells[colIndex1];
      var cell2 = row.cells[colIndex2];
      var siblingCell1 = row.cells[Number(colIndex1) + 1];
      row.insertBefore(cell1, cell2);
      row.insertBefore(cell2, siblingCell1);
    }
  }
}

var tableToExcel = (function() {
  var uri = 'data:application/vnd.ms-excel;base64,'
    , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
    , base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) }
    , format = function(s, c) { return s.replace(/{(\w+)}/g, function(m, p) { return c[p]; }) }
  return function(table, name) {
    if (!table.nodeType) table = document.getElementById(table)
    var ctx = {worksheet: name || 'Worksheet', table: table.innerHTML}
    window.location.href = uri + base64(format(template, ctx))
  }
})()