/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Connect database lấy dữ liệu

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Thử xử lý form select của Bootstrap
    $("option[selected]").hide()
    // $("option[selected]").removeAttr("selected")
    $("select.taskpane__btn").on("change",function() {
      let type = $(this).val()
        createReport(type)
    })

    

    // create template report
    // $("#create-report").click(() => { createReport() })


    //clear table
    document.getElementById("clear-table").onclick = clearTable;

    // connect database
    document.getElementById('connect-db').onclick = connectDB;

    // Open Dialog
    document.getElementById("open-dialog").onclick = openDialog;
  }
});
// Hàm tạo template báo cáo
async function createReport(type) {
  if (type == 1) {
    await clearTable()
    await Excel.run(async (context) => {
      //lấy worksheet đang làm việc hiện tại
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      //xóa đường grid lines
      // currentWorksheet.showGridlines = false

      //tên công ty
      let nameCompanyRange = currentWorksheet.getRange("A1:B1")
      //quốc hiệu
      let nationalNameRange = currentWorksheet.getRange("E1:H1")
      //tiêu ngữ
      let crestRange = currentWorksheet.getRange("E2:H2")
      //tên báo cáo
      let reportNameRange = currentWorksheet.getRange("B5:G6")
      //date
      let dateRange = currentWorksheet.getRange("A2:C2")

      let date = new Date()
      let dateValue = `Ngày ${date.getDate()} Tháng ${date.getMonth()} Năm ${date.getFullYear()}`



      nationalNameRange.merge()
      crestRange.merge()
      nameCompanyRange.merge()
      reportNameRange.merge()
      dateRange.merge()

      let nameCompany = currentWorksheet.getRange("A1")
      let nationalName = currentWorksheet.getRange("E1")
      let crest = currentWorksheet.getRange("E2")
      let reportName = currentWorksheet.getRange("B5")


      currentWorksheet.getRange("A2").values = [[`${dateValue}`]]

      nameCompany.values = [["Công ty Rạng Đông"]]
      nationalName.values = [['Cộng hòa xã hội chủ nghĩa Việt Nam']]
      crest.values = [['Độc lập - Tự do - Hạnh phúc']]
      reportName.values = [['BÁO CÁO SẢN XUẤT CÔNG NHÂN']]

      nationalName.format.horizontalAlignment = "Center"
      crest.format.horizontalAlignment = "Center"
      nameCompany.format.horizontalAlignment = "Left"
      reportName.format.horizontalAlignment = "Center"
      reportName.format.font.bold = true
      reportName.format.font.size = 18

      let nameWorker = currentWorksheet.getRange("B9")
      let productPlant = currentWorksheet.getRange("F9")
      let timeStart = currentWorksheet.getRange("B10")
      let timeEnd = currentWorksheet.getRange("D10")
      let employeeCode = currentWorksheet.getRange("D9")
      let reporter = currentWorksheet.getRange("F19")

      nameWorker.values = [["Họ và tên:"]]
      productPlant.values = [['Phân xưởng:']]
      timeStart.values = [['Từ ngày:']]
      timeEnd.values = [['Đến ngày:']]
      employeeCode.values = [['Mã nhân viên:']]
      reporter.values = [["Người lập báo cáo"]]

      nameWorker.format.horizontalAlignment = "Left"
      nameWorker.format.font.size = 12
      nameWorker.format.font.italic = true

      reporter.format.horizontalAlignment = "Left"
      reporter.format.font.size = 12
      reporter.format.font.italic = true

      employeeCode.format.horizontalAlignment = "Left"
      employeeCode.format.font.size = 12
      employeeCode.format.font.italic = true

      productPlant.format.horizontalAlignment = "Left"
      productPlant.format.font.size = 12
      productPlant.format.font.italic = true

      timeStart.format.horizontalAlignment = "Left"
      timeStart.format.font.size = 12
      timeStart.format.font.italic = true

      timeEnd.format.horizontalAlignment = "Left"
      timeEnd.format.font.size = 12
      timeEnd.format.font.italic = true






      //tạo bảng mới có header sau khi nhân được kết quả
      const expensesTable = currentWorksheet.tables.add("B12:G12", true /*hasHeaders*/);
      expensesTable.name = "ReportData";
      // // TODO2: Queue commands to populate the table with data.
      expensesTable.getHeaderRowRange().values =
        [["STT", "Mã Công Việc", "Tên Công Việc", "Tổng chỉ tiêu", "Sản phẩm đạt", "Sản phẩm lỗi"]];


      expensesTable.getHeaderRowRange().format.fill.color = "#009879";
      expensesTable.getDataBodyRange().format.fill.color = "#FFFFFF";
      // expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
      // expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        currentWorksheet.getUsedRange().format.autofitColumns();
        currentWorksheet.getUsedRange().format.autofitRows();
        currentWorksheet.getUsedRange().format.horizontalAlignment = "Center"
      }

      currentWorksheet.getRange("A2").format.horizontalAlignment = "Left"
      currentWorksheet.getRange("A2").format.font.italic = true
      await context.sync();
    })
      .catch((error) => {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
  } else if (type == 2) {
    await clearTable()

    await Excel.run(async (context) => {
      //lấy worksheet đang làm việc hiện tại
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      //xóa đường grid lines
      // currentWorksheet.showGridlines = false

      //tên công ty
      let nameCompanyRange = currentWorksheet.getRange("A1:B1")
      //quốc hiệu
      let nationalNameRange = currentWorksheet.getRange("E1:H1")
      //tiêu ngữ
      let crestRange = currentWorksheet.getRange("E2:H2")
      //tên báo cáo
      let reportNameRange = currentWorksheet.getRange("B5:G6")
      //date
      let dateRange = currentWorksheet.getRange("A2:C2")

      let date = new Date()
      let dateValue = `Ngày ${date.getDate()} Tháng ${date.getMonth()} Năm ${date.getFullYear()}`



      nationalNameRange.merge()
      crestRange.merge()
      nameCompanyRange.merge()
      reportNameRange.merge()
      dateRange.merge()

      let nameCompany = currentWorksheet.getRange("A1")
      let nationalName = currentWorksheet.getRange("E1")
      let crest = currentWorksheet.getRange("E2")
      let reportName = currentWorksheet.getRange("B5")


      currentWorksheet.getRange("A2").values = [[`${dateValue}`]]

      nameCompany.values = [["Công ty Rạng Đông"]]
      nationalName.values = [['Cộng hòa xã hội chủ nghĩa Việt Nam']]
      crest.values = [['Độc lập - Tự do - Hạnh phúc']]
      reportName.values = [['BÁO CÁO SẢN XUẤT PHÂN XƯỞNG']]

      nationalName.format.horizontalAlignment = "Center"
      crest.format.horizontalAlignment = "Center"
      nameCompany.format.horizontalAlignment = "Left"
      reportName.format.horizontalAlignment = "Center"
      reportName.format.font.bold = true
      reportName.format.font.size = 18

      let nameWorker = currentWorksheet.getRange("B9")

      let timeStart = currentWorksheet.getRange("B10")
      let timeEnd = currentWorksheet.getRange("D10")
      let reporter = currentWorksheet.getRange("F19")

      nameWorker.values = [["Phân xưởng"]]

      timeStart.values = [['Từ ngày:']]
      timeEnd.values = [['Đến ngày:']]
      reporter.values = [["Người lập báo cáo"]]

      nameWorker.format.horizontalAlignment = "Left"
      nameWorker.format.font.size = 12
      nameWorker.format.font.italic = true

      reporter.format.horizontalAlignment = "Left"
      reporter.format.font.size = 12
      reporter.format.font.italic = true



      timeStart.format.horizontalAlignment = "Left"
      timeStart.format.font.size = 12
      timeStart.format.font.italic = true

      timeEnd.format.horizontalAlignment = "Left"
      timeEnd.format.font.size = 12
      timeEnd.format.font.italic = true






      //tạo bảng mới có header sau khi nhân được kết quả
      const expensesTable = currentWorksheet.tables.add("B12:G12", true /*hasHeaders*/);
      expensesTable.name = "ReportData";
      // // TODO2: Queue commands to populate the table with data.
      expensesTable.getHeaderRowRange().values =
        [["STT", "Mã Công Việc", "Tên Công Việc", "Tổng chỉ tiêu", "Sản phẩm đạt", "Sản phẩm lỗi"]];

      expensesTable.getHeaderRowRange().format.fill.color = "#009879";
      expensesTable.getDataBodyRange().format.fill.color = "#FFFFFF";

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        currentWorksheet.getUsedRange().format.autofitColumns();
        currentWorksheet.getUsedRange().format.autofitRows();
        currentWorksheet.getUsedRange().format.horizontalAlignment = "Center"
      }

      currentWorksheet.getRange("A2").format.horizontalAlignment = "Left"
      currentWorksheet.getRange("A2").format.font.italic = true
      await context.sync();
    })
      .catch((error) => {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
  } else {
    await clearTable()

    await Excel.run(async (context) => {
      //lấy worksheet đang làm việc hiện tại
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      //xóa đường grid lines
      // currentWorksheet.showGridlines = false

      //tên công ty
      let nameCompanyRange = currentWorksheet.getRange("A1:B1")
      //quốc hiệu
      let nationalNameRange = currentWorksheet.getRange("E1:H1")
      //tiêu ngữ
      let crestRange = currentWorksheet.getRange("E2:H2")
      //tên báo cáo
      let reportNameRange = currentWorksheet.getRange("B5:G6")
      //date
      let dateRange = currentWorksheet.getRange("A2:C2")

      let date = new Date()
      let dateValue = `Ngày ${date.getDate()} Tháng ${date.getMonth()} Năm ${date.getFullYear()}`



      nationalNameRange.merge()
      crestRange.merge()
      nameCompanyRange.merge()
      reportNameRange.merge()
      dateRange.merge()

      let nameCompany = currentWorksheet.getRange("A1")
      let nationalName = currentWorksheet.getRange("E1")
      let crest = currentWorksheet.getRange("E2")
      let reportName = currentWorksheet.getRange("B5")


      currentWorksheet.getRange("A2").values = [[`${dateValue}`]]

      nameCompany.values = [["Công ty Rạng Đông"]]
      nationalName.values = [['Cộng hòa xã hội chủ nghĩa Việt Nam']]
      crest.values = [['Độc lập - Tự do - Hạnh phúc']]
      reportName.values = [['BÁO CÁO TIẾN ĐỘ SẢN XUẤT']]

      nationalName.format.horizontalAlignment = "Center"
      crest.format.horizontalAlignment = "Center"
      nameCompany.format.horizontalAlignment = "Left"
      reportName.format.horizontalAlignment = "Center"
      reportName.format.font.bold = true
      reportName.format.font.size = 18

      let nameWorker = currentWorksheet.getRange("B9")
      let codePX = currentWorksheet.getRange("D9")

      let timeStart = currentWorksheet.getRange("B12")
      let timeEnd = currentWorksheet.getRange("D12")
      let reporter = currentWorksheet.getRange("D19")
      let target = currentWorksheet.getRange("F12")
      let task = currentWorksheet.getRange("B11")
      let code_task = currentWorksheet.getRange("D11")

      nameWorker.values = [["Phân xưởng:"]]
      codePX.values = [["Mã PX:"]]
      target.values = [["Chỉ tiêu SX:"]]
      task.values = [["Công việc:"]]
      code_task.values = [["Mã SX:"]]

      timeStart.values = [['Từ ngày:']]
      timeEnd.values = [['Đến ngày:']]
      reporter.values = [["Người lập báo cáo"]]

      task.format.horizontalAlignment = "Left"
      task.format.font.size = 12
      task.format.font.italic = true

      code_task.format.horizontalAlignment = "Left"
      code_task.format.font.size = 12
      code_task.format.font.italic = true

      codePX.format.horizontalAlignment = "Left"
      codePX.format.font.size = 12
      codePX.format.font.italic = true

      target.format.horizontalAlignment = "Left"
      target.format.font.size = 12
      target.format.font.italic = true

      nameWorker.format.horizontalAlignment = "Left"
      nameWorker.format.font.size = 12
      nameWorker.format.font.italic = true

      reporter.format.horizontalAlignment = "Left"
      reporter.format.font.size = 12
      reporter.format.font.italic = true



      timeStart.format.horizontalAlignment = "Left"
      timeStart.format.font.size = 12
      timeStart.format.font.italic = true

      timeEnd.format.horizontalAlignment = "Left"
      timeEnd.format.font.size = 12
      timeEnd.format.font.italic = true






      //tạo bảng mới có header sau khi nhân được kết quả
      const expensesTable = currentWorksheet.tables.add("B15:E15", true /*hasHeaders*/);
      expensesTable.name = "ReportData";
      // // TODO2: Queue commands to populate the table with data.
      expensesTable.getHeaderRowRange().values =
        [["Ngày", "Chỉ tiêu", "Tổng SP đạt", "Tổng SP lỗi"]];

      expensesTable.getHeaderRowRange().format.fill.color = "#009879";
      expensesTable.getDataBodyRange().format.fill.color = "#FFFFFF";

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        currentWorksheet.getUsedRange().format.autofitColumns();
        currentWorksheet.getUsedRange().format.autofitRows();
        currentWorksheet.getUsedRange().format.horizontalAlignment = "Center"
      }

      currentWorksheet.getRange("A2").format.horizontalAlignment = "Left"
      currentWorksheet.getRange("A2").format.font.italic = true
      await context.sync();
    })
      .catch((error) => {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
  }
}

// Hàm lấy dữ liệu từ database
async function connectDB() {
  await Excel.run(async (context) => {
    const id_worker = document.getElementById("id_worker_taskpane").innerHTML;
    const id_PX = document.getElementById("abc").innerHTML;
    const time_work = document.getElementById("time_taskpane").innerHTML;
    const type_report = document.getElementById("user-name").innerHTML;
    const id_task = document.getElementById("task_taskpane").innerHTML;

    if (!type_report) {
      document.querySelector(".taskpane__error").innerHTML = "Bạn cần nhập đủ thông tin"
      return
    } else {
      if (type_report == "Báo cáo sản xuất công nhân") {
        if (!id_worker || !time_work) {
          document.querySelector(".taskpane__error").innerHTML = "Bạn cần nhập đủ thông tin"
          return
        }
      } else if (type_report == "Báo cáo sản xuất phân xưởng") {
        if (!id_PX || !time_work) {
          document.querySelector(".taskpane__error").innerHTML = "Bạn cần nhập đủ thông tin"
          return
        }
      } else {
        if (!id_PX || !id_task) {
          document.querySelector(".taskpane__error").innerHTML = "Bạn cần nhập đủ thông tin"
          return
        }
      }
    }
    document.querySelector(".taskpane__error").innerHTML = ""

    //scan worksheet để lấy dữ liệu:
    //lấy worksheet đang làm việc hiện tại
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

    //mảng để lưu các giá trị người dùng nhập vào
    let dataClient = [];


    let ranges = currentWorksheet.getUsedRange();
    // console.log(ranges)
    //cái sau là bất đồng bộ nên cần chờ
    ranges.load("values")
    // ranges.load("address")
    await context.sync();

    // const RegExp = /^[\{]([a-z]|\-|\_)+[\}]$/
    const RegExp = /^[\{]([a-z]|[0-9]|\(|\)|\-|\_|\/)+[\}]$/
    // console.log(ranges.address)
    for (let value of ranges.values) {
      for (let text of value) {
        if (RegExp.test(text)) {
          dataClient.push(text)
        }
      }
    }

    //mảng kết quả trả về
    let result = []

    if (type_report == "Báo cáo sản xuất công nhân") {
      const options = {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(
          {
            worker: id_worker,
            time: time_work,
            dataClient
          }
        )
      };
      await fetch('http://localhost:8080/data', options)
        .then((res) => res.text())
        .then(data => {
          console.log("Giá trị Server trả về là : " + data)
          result.push(JSON.parse(data))
        }
        );
    } else if (type_report == "Báo cáo sản xuất phân xưởng") {
      const options = {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(
          {
            productplant: id_PX,
            time: time_work,
            dataClient
          }
        )
      };
      await fetch('http://localhost:8080/data', options)
        .then((res) => res.text())
        .then(data => {
          console.log("Giá trị Server trả về là : " + data)
          result.push(JSON.parse(data))
        }
        );
    } else {
      const options = {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(
          {
            productplant: id_PX,
            task: id_task,
            dataClient
          }
        )
      };
      await fetch('http://localhost:8080/data', options)
        .then((res) => res.text())
        .then(data => {
          console.log("Giá trị Server trả về là : " + data)
          result.push(JSON.parse(data))
        }
        );
    }
    
    console.log("Excel nhận được sau khi fetch xong: " + result)
    // const RegExp2 = /^[\{]([a-z]|\-|\_)+[\}]$/

    // for (let value of ranges.values) {
    //   for (let text of value) {
    //     if (RegExp2.test(text)) {
    //       // console.log(text)
    //     }
    //   }
    // }

    //đổ dữ liệu từ server ra file excel:
    for (let i = 0; i < result[0].length; i++) {
      let foundRange = ranges.find(`{${result[0][i].data}}`, {
        completeMatch: true, // Match the whole cell value.
        matchCase: false, // Don't match case.
        searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
      });
      foundRange.load("address");
      await context.sync();
      // console.log(`Vị trí của ${result[i].data} là ` + foundRange.address);
      //thay đổi vị trí từ Food thành Kiền
      if (foundRange.address) {
        foundRange.values = [[`${result[0][i].value}`]]
        foundRange.format.autofitColumns();
        await context.sync();
      } else {
        break
      }
    }
  })
    .catch((error) => {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}


// Hàm xóa toàn bộ worksheet
async function clearTable() {
  await Excel.run(async (context) => {

    // TODO1: Queue table creation logic here.

    //lấy worksheet đang làm việc hiện tại
    $('select.taskpane__btn').prop('selectedIndex',0);
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

    currentWorksheet.getRange().clear()
    await context.sync();
  })
    .catch((error) => {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}


// Xử lý các dialog
let dialog = null;

function openDialog() {
  // TODO1: Call the Office Common API that opens a dialog
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html',
    { height: 32, width: 30 },

    // TODO2: Add callback parameter.
    function (result) {
      // console.log(result)
      dialog = result.value;
      // console.log(result.value)

      // Nhận được loại báo cáo
      // Báo cáo cá nhân: value = person
      // Báo cáo theo xưởng: value = productplant
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processTypeReport);
    }

  );
}

function processTypeReport(arg) {
  // console.log(arg.message)
  if (arg.message == "worker") {
    document.getElementById("user-name").innerHTML = "Báo cáo sản xuất công nhân";
  } else if (arg.message == "productplant") {
    document.getElementById("user-name").innerHTML = "Báo cáo sản xuất phân xưởng";
  } else {
    document.getElementById("user-name").innerHTML = "Báo cáo tiến độ sản xuất";
  }

  // console.log("2")
  dialog.close();
  // Mở dialog thứ 2 để chọn xưởng
  // if (arg.message == "worker") {
  //   openProductPlantDialog();
  // } 
  openProductPlantDialog();

}

function openProductPlantDialog() {
  // console.log("3")
  Office.context.ui.displayDialogAsync("https://localhost:3000/ProductPlant.html", { width: 35, height: 30 },
    (result) => {

      // console.log(result)

      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          // console.log("Dialog 1 chưa đóng hẳn")
          openProductPlantDialog(); // Recursive call
        }
        else {
          // Handle other errors
          console.log("Lỗi gì đó")
          openProductPlantDialog(); // Recursive call

        }
      }
      console.log("Mở dialog 2 thành công")
      dialog = result.value;
      // console.log(result.value)
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processProductPlantValue);
      // console.log("Lấy dữ liệu từ dialog")
    }
  );
}

function processProductPlantValue(arg) {
  // Gửi id_productplant lên taskpane
  document.getElementById("abc").innerHTML = arg.message;

  dialog.close();

  let type_report = document.getElementById("user-name").innerHTML
  if (type_report == "Báo cáo sản xuất công nhân") {
    console.log("else 1")
    openWorkerDialog()
  } else if (type_report == "Báo cáo sản xuất phân xưởng") {
    console.log("Chuẩn bị mở dialog lấy thời gian")
    // document.querySelector(".taskpane__error").innerHTML = "test"
    openDateDialog()
  } else {
    openTaskDialog()
  }
}

function openWorkerDialog() {
  let id_productplant = document.getElementById("abc").innerHTML
  Office.context.ui.displayDialogAsync(`https://localhost:3000/Worker.html?id=${id_productplant}`, { width: 30, height: 36 },
    (result) => {
      // console.log(result)
      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          // console.log("Dialog 1 chưa đóng hẳn")
          openWorkerDialog(); // Recursive call
        }
        else {
          // Handle other errors
          console.log("Lỗi gì đó ở Dialog Worker")
          openWorkerDialog(); // Recursive call
        }
      }
      console.log("Mở dialog 3 thành công")

      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processWorkerValue);
      console.log("Nhận event Dialog Worker")

    }
  );
}

function processWorkerValue(arg) {
  console.log("test hàm dialog 3")
  document.getElementById("id_worker_taskpane").innerHTML = arg.message
  dialog.close()

  // Mở Time Dialog
  openDateDialog()
}

function openDateDialog() {
  Office.context.ui.displayDialogAsync("https://localhost:3000/Date.html?", { width: 40, height: 40 },
    (result) => {

      console.log("Chạy vào hàm mở dialog 4")
      // console.log(result)
      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          // console.log("Dialog 1 chưa đóng hẳn")
          openDateDialog(); // Recursive call
        }
        else {
          // Handle other errors
          console.log("Lỗi gì đó ở Dialog Time")
          openDateDialog(); // Recursive call
        }
      }
      console.log("Mở dialog 4 thành công")
      dialog = result.value;
      // console.log(result.value)
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processTimeValue);
      // console.log("Lấy dữ liệu từ dialog")
    }
  );
}

function processTimeValue(arg) {
  console.log("test hàm dialog 4")
  console.log(arg.message)
  document.getElementById("time_taskpane").innerHTML = arg.message
  dialog.close()

  // Gọi hàm tạo template báo cáo ở đây
  // createReport()
}

function openTaskDialog() {
  let id_productplant = document.getElementById("abc").innerHTML
  Office.context.ui.displayDialogAsync(`https://localhost:3000/task.html?id=${id_productplant}`, { width: 30, height: 36 },
    (result) => {
      // console.log(result)
      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          // console.log("Dialog 1 chưa đóng hẳn")
          openTaskDialog(); // Recursive call
        }
        else {
          // Handle other errors
          console.log("Lỗi gì đó ở Dialog Worker")
          openTaskDialog(); // Recursive call
        }
      }

      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processTaskValue);
      console.log("Nhận event Dialog Worker")

    }
  );
}

function processTaskValue(arg) {
  console.log(arg.message)
  document.getElementById("task_taskpane").innerHTML = arg.message
  dialog.close()
}


