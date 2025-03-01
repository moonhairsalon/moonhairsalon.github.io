function createOrUpdateSheetLuong() {
    var ui = SpreadsheetApp.getUi(); // Lấy đối tượng UI

    // sheet cấu hình
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Hiển thị hộp thoại xác nhận
    var response = ui.alert("Xác nhận", "Bạn có muốn tạo bảng lương cho tháng " + sheet.getRange("A1").getValue() + "?", ui.ButtonSet.YES_NO);

    // Kiểm tra phản hồi của người dùng
    if (response == ui.Button.YES) {
        var lastColumn = sheet.getLastColumn();
        if (lastColumn > 17) {
            var rangeUnMerge = sheet.getRange(3, 15, 2, lastColumn);
            rangeUnMerge.breakApart();
            rangeUnMerge.clearContent().clearFormat().setBackground("#ffffff").setBorder(false, false, false, false, false, false).setFontFamily("Arial").setFontSize(12);
        }

        sheet.setColumnWidth(15, 120)
            .setColumnWidth(16, 150)
            .setColumnWidth(17, 130)
            .setColumnWidth(18, 130)
            .setColumnWidth(19, 150)
            .setColumnWidth(20, 150)
            .setColumnWidth(21, 150)
            .setColumnWidth(22, 150)
            .setColumnWidth(23, 150)
            .setColumnWidth(24, 150)
            .setColumnWidth(25, 150)
            .setColumnWidth(26, 150);
        const cellTitle = sheet.getRange("o1");
        let title = "Lương " + sheet.getRange("A1").getValue();
        cellTitle.setValue(title);
        cellTitle.setFontWeight("bold");
        cellTitle.setFontSize(17).setHorizontalAlignment("left");
        var values = [["Ngày", "Tên khách", "Tiền bill", "Tổng bill ngày"]];

        const headerCommons = sheet.getRange(3, 15, 2, 4);
        for (var col = 15; col <= 18; col++) {
            // Gộp ô tại hàng 3 và hàng 4 cho mỗi cột
            sheet.getRange(3, col, 2, 1).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
            sheet.getRange(3, col).setValue(values[0][col - 15]);
            sheet.getRange(3, col).setVerticalAlignment("middle").setHorizontalAlignment("center");
        }
        headerCommons.setFontWeight("bold");
        headerCommons.setFontSize(12);
        headerCommons.setBackground("#d3d3d3");
        headerCommons.setBorder(true, true, true, true, true, true); // Đặt đường viền cho các cạnh trên, dưới, trái, phải

        let lastRowThoChinh = sheet.getRange(sheet.getMaxRows(), 10).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
        let lastRowThoPhu = sheet.getRange(sheet.getMaxRows(), 12).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

        let colThoChinh = sheet.getRange("J4:J" + lastRowThoChinh);
        let thoChinh = colThoChinh.getValues();

        let colThoPhu = sheet.getRange("L4:L" + lastRowThoPhu);
        let thoPhu = colThoPhu.getValues();

        var startCellTho = sheet.getRange("S4"); // Lấy ô E4
        let coutTho = 0;

        for (let i = 0; i < thoChinh.length; i++) {
            startCellTho.offset(0, coutTho).setValue(thoChinh[i])
                .setFontWeight("bold")
                .setFontSize(12)
                .setBackground(colThoChinh.getCell(i + 1, 1).getBackground())
                .setBorder(true, true, true, true, true, true)
                .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
            coutTho += 1;
        }
        sheet.getRange(3, 19, 1, thoChinh.length).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
        sheet.getRange(3, 19).setValue("Thợ chính")
            .setFontWeight("bold")
            .setFontSize(12)
            .setBackground("#d3d3d3")
            .setBorder(true, true, true, true, true, true)
            .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang

        for (let i = 0; i < thoPhu.length; i++) {
            startCellTho.offset(0, coutTho).setValue(thoPhu[i])
                .setFontWeight("bold")
                .setFontSize(12)
                .setBackground(colThoPhu.getCell(i + 1, 1).getBackground())
                .setBorder(true, true, true, true, true, true)
                .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
            coutTho += 1;
        }
        sheet.getRange(3, 19 + thoChinh.length, 1, thoPhu.length).merge(); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
        sheet.getRange(3, 19 + thoChinh.length).setValue("Thợ phụ")
            .setFontWeight("bold")
            .setFontSize(12)
            .setBackground("#d3d3d3")
            .setBorder(true, true, true, true, true, true)
            .setHorizontalAlignment("center"); // Căn giữa theo chiều ngang
        sheet.getRange(3, 19 + thoChinh.length + thoPhu.length, 2, 1).merge()
            .setValue("Ngày sửa")
            .setFontWeight("bold")
            .setFontSize(12)
            .setBackground("#d3d3d3")
            .setBorder(true, true, true, true, true, true)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        // ui.alert("Tạo bảng lương thành công.");
    } else {
    }

}

function tinhLuong() {
    var ui = SpreadsheetApp.getUi(); // Lấy đối tượng UI

    let dongBatDauLuong = 5;
    let dongBatDauDT = 5;
    let startColTho = 19;
    let tongDoanhThu = 0;
    let cotBatDauLuong = 15;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRowSheet = sheet.getLastRow(); // Lấy dòng cuối cùng có dữ liệu của toàn bộ sheet

    // Hiển thị hộp thoại xác nhận
    var response = ui.alert("Xác nhận", "Bạn có muốn tính lương cho tháng " + sheet.getRange("A1").getValue() + "?", ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {

        // clear data + format
        var dataGetLastRowLuong = sheet.getRange("Q5:Q" + lastRowSheet).getValues(); // Lấy dữ liệu cột Q
        var lastRowLuong = 4;

        for (var i = 0; i < dataGetLastRowLuong.length; i++) {
            if (dataGetLastRowLuong[i] && dataGetLastRowLuong[i][0] !== "") {
                lastRowLuong = lastRowLuong + 1; // Đếm số dòng có dữ liệu
            }
        }

        var dataGetLastRow = sheet.getRange("A4:A" + lastRowSheet).getValues(); // Lấy dữ liệu cột A
        var lastRow = dongBatDauDT - 1;
        for (var i = 0; i < dataGetLastRow.length; i++) {
            if (dataGetLastRow[i] && dataGetLastRow[i][0] !== "") {
                lastRow = lastRow + 1; // Đếm số dòng có dữ liệu
            }
        }

        Logger.log('Dong cuoi cung ' + lastRow)

        if (lastRowLuong > 8) { // nếu đã fill dữ liệu lương
            // clear dòng tổng kết
            sheet.getRange(lastRowLuong, 15, 2, 20).clearContent().clearFormat().setBackground("#ffffff").setBorder(false, false, false, false, false, false).setFontFamily("Arial").setFontSize(12).setHorizontalAlignment("center");
            // clear merge của cột ngày
            sheet.getRange(dongBatDauLuong, 15, lastRowLuong, 1).breakApart();
            sheet.getRange(dongBatDauLuong, 18, lastRowLuong, 1).breakApart();
        }

        // Doanh thu - Ngày
        const columnDateDT = sheet.getRange("A" + dongBatDauDT + ":A" + lastRow);
        const dateDT = columnDateDT.getValues();
        // Doanh thu - tên khách
        const columnCustomerDT = sheet.getRange("B" + dongBatDauDT + ":B" + lastRow);
        const customerDT = columnCustomerDT.getValues();
        // Doanh thu - tiền bill
        const columnBillDT = sheet.getRange("C" + dongBatDauDT + ":C" + lastRow);
        const billDT = columnBillDT.getValues();
        // Doanh thu - thợ chính
        const columnThoChinhDT = sheet.getRange("E" + dongBatDauDT + ":E" + lastRow);
        const thoChinhDT = columnThoChinhDT.getValues();
        // Doanh thu - thợ phụ
        const columnThoPhuDT = sheet.getRange("F" + dongBatDauDT + ":F" + lastRow);
        const thoPhuDT = columnThoPhuDT.getValues();

        // Doanh thu - trạng thái tính lương
        const columnStatusDT = sheet.getRange("H" + dongBatDauDT + ":H" + lastRow);
        const statusDT = columnStatusDT.getValues();

        // Lương - ngày
        const columnDateL = sheet.getRange("O" + dongBatDauLuong + ":O" + lastRow + 1);
        const dateL = columnDateL.getValues();
        // Lương - tên khách
        const columnCustomerL = sheet.getRange("P" + dongBatDauLuong + ":P" + lastRow + 1);
        const customerL = columnCustomerL.getValues();
        // Lương - tiền bill
        const columnBillL = sheet.getRange("Q" + dongBatDauLuong + ":Q" + lastRow + 1);
        columnBillL.setNumberFormat("#,##0");
        const billL = columnBillL.getValues();

        // Danh sách thợ
        let listThoChinh = new Array();
        let listThoPhu = new Array();

        let lastRowThoChinh = sheet.getRange(sheet.getMaxRows(), 10).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
        let lastRowThoPhu = sheet.getRange(sheet.getMaxRows(), 12).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

        let colThoChinh = sheet.getRange("J4:J" + lastRowThoChinh);
        let thoChinh = colThoChinh.getValues();
        let colLuongThoChinh = sheet.getRange("K4:K" + lastRowThoChinh);
        let luongThoChinh = colLuongThoChinh.getValues();

        let colThoPhu = sheet.getRange("L4:L" + lastRowThoPhu);
        let thoPhu = colThoPhu.getValues();
        let colLuongThoPhu = sheet.getRange("M4:M" + lastRowThoPhu);
        let luongThoPhu = colLuongThoPhu.getValues();

        var headerThoChinh = sheet.getRange(4, 19, 1, thoChinh.length).getValues()[0];
        var headerThoPhu = sheet.getRange(4, 19 + thoChinh.length, 1, thoPhu.length).getValues()[0];

        // start lấy thông tin thợ chính
        for (let i = 0; i < thoChinh.length; i++) {
            listThoChinh.push({
                name: thoChinh[i][0],
                luong: luongThoChinh[i][0],
                index: headerThoChinh.indexOf(thoChinh[i][0]) + startColTho,
                color: colThoChinh.getCell(i + 1, 1).getBackground(),
                tongLuong: 0
            })
        }
        // end lấy thông tin thợ chính
        // start lấy thông tin thợ phụ
        for (let i = 0; i < thoPhu.length; i++) {
            listThoPhu.push({
                name: thoPhu[i][0],
                luong: luongThoPhu[i][0],
                index: headerThoPhu.indexOf(thoPhu[i][0]) + startColTho + thoChinh.length,
                color: colThoPhu.getCell(i + 1, 1).getBackground(),
                tongLuong: 0
            })
        }
        // end lấy thông tin thợ phụ
        //start fill data
        for (let i = 0; i < dateDT.length; i++) {
            if (dateDT[i] != undefined && dateDT[i] != "" && statusDT[i][0] == 0) {

                // reset value dòng
                sheet.getRange(dongBatDauLuong + i, 15, 1, 20).clearContent().clearFormat().setBackground("#ffffff").setBorder(false, false, false, false, false, false).setFontFamily("Arial").setFontSize(12);

                dateL[i] = dateDT[i];
                columnDateL.getCell(i + 1, 1).setFontWeight("bold");
                customerL[i] = customerDT[i];
                billDT[i][0] = billDT[i][0] * 1000;
                billL[i] = billDT[i];

                const curentThoChinh = listThoChinh.find(item => item.name == thoChinhDT[i]);
                if (curentThoChinh) {
                    sheet.getRange(dongBatDauLuong + i, curentThoChinh.index)
                        .setValue(billDT[i] / 100 * curentThoChinh.luong)
                        .setFontSize("12")
                        .setBackground(curentThoChinh.color)
                        .setNumberFormat("#,##0");
                }

                const curentThoPhu = listThoPhu.find(item => item.name == thoPhuDT[i]);
                if (curentThoPhu) {
                    if (customerDT[i][0].toLowerCase() == "bsp" || customerDT[i][0].toLowerCase() == "gội" || customerDT[i][0].toLowerCase() == "cắt") {
                        sheet.getRange(dongBatDauLuong + i, curentThoPhu.index)
                            .setValue(billDT[i] / 100 * 20)
                            .setFontSize("12")
                            .setBackground(curentThoPhu.color)
                            .setNumberFormat("#,##0");
                    } else {
                        sheet.getRange(dongBatDauLuong + i, curentThoPhu.index)
                            .setValue(billDT[i] / 100 * curentThoPhu.luong)
                            .setFontSize("12")
                            .setBackground(curentThoPhu.color)
                            .setNumberFormat("#,##0");
                    }
                }
                var currentDate = new Date(); // Lấy ngày hiện tại
                sheet.getRange(dongBatDauLuong + i, startColTho + listThoPhu.length + listThoChinh.length).setValue(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "HH:MM dd/MM/yyyy")).setFontSize("12");
                statusDT[i] = [1];
            }
        }

        // end fill data
        // set data
        columnDateL.setValues(dateL).setFontSize("12");
        columnCustomerL.setValues(customerL).setFontSize("12");
        columnBillL.setValues(billL).setFontSize("12");
        columnStatusDT.setValues(statusDT);
        // end set data

        // start tinh tong luong
        for (let i = 0; i < listThoChinh.length; i++) {
            listThoChinh[i].tongLuong = 0;
            var listLuong = sheet.getRange(5, listThoChinh[i].index, lastRow, 1).getValues();
            if (listLuong) {
                for (let l = 0; l < listLuong.length; l++) {
                    if (listLuong[l]) {
                        listThoChinh[i].tongLuong += Number(listLuong[l]);
                    }
                }
            }
        }
        for (let i = 0; i < listThoPhu.length; i++) {
            listThoPhu[i].tongLuong = 0;
            var listLuong = sheet.getRange(5, listThoPhu[i].index, lastRow, 1).getValues();
            if (listLuong) {
                for (let l = 0; l < listLuong.length; l++) {
                    if (listLuong[l]) {
                        listThoPhu[i].tongLuong += Number(listLuong[l]);
                    }
                }
            }
        }
        var listDoanhThu = sheet.getRange(5, 17, lastRow, 1).getValues();
        if (listDoanhThu) {
            for (let l = 0; l < listDoanhThu.length; l++) {
                if (listDoanhThu[l]) {
                    tongDoanhThu += Number(listDoanhThu[l]);
                }
            }
        }
        // end tinh tong luong
        // ô Tổng
        sheet.getRange(lastRow + 1, 15, 2, 2).merge()
            .setValue("Tổng").setFontSize("14")
            .setFontWeight("bold");

        // ô Tổng lương số
        sheet.getRange(lastRow + 1, 17, 2, 1).merge()
            .setValue(tongDoanhThu).setFontSize("14")
            .setFontWeight("bold")
            .setNumberFormat("#,##0");
        // ô lương chữ
        sheet.getRange(lastRow + 1, 18, 2, 1).merge()
            .setValue("Lương").setFontSize("14")
            .setFontWeight("bold");

        // in tổng lương thợ chính
        for (let i = 0; i < listThoChinh.length; i++) {
            sheet.getRange(lastRow + 1, listThoChinh[i].index, 2, 1).merge()
                .setValue(listThoChinh[i].tongLuong).setFontSize("14")
                .setFontWeight("bold")
                .setBackground(listThoChinh[i].color)
                .setNumberFormat("#,##0");
        }
        // in tổng lương thợ phụ
        for (let i = 0; i < listThoPhu.length; i++) {
            sheet.getRange(lastRow + 1, listThoPhu[i].index, 2, 1).merge()
                .setValue(listThoPhu[i].tongLuong).setFontSize("14")
                .setFontWeight("bold")
                .setBackground(listThoPhu[i].color)
                .setNumberFormat("#,##0");
        }

        // ô trống cuối
        sheet.getRange(lastRow + 1, +startColTho + listThoPhu.length + listThoChinh.length, 2, 1).merge()
            .setFontWeight("bold");

        // start merge cột ngày giống nhau
        var startRow = 5; // Dòng bắt đầu từ A5

        // Lấy tất cả giá trị trong cột O từ dòng 5 trở đi
        var data = sheet.getRange("O" + dongBatDauLuong + ":O" + (lastRow)).getValues();

        var startMergeRow = startRow;  // Dòng bắt đầu merge
        let curentDateCheck = data[0][0]
        let coutSame = 1;
        let currentColor = 1;
        let hangColor = "";
        Logger.log('So data quét' + data.length)
        for (var i = 1; i <= data.length; i++) {
            let curentDateCheckStr = curentDateCheck ? Utilities.formatDate(curentDateCheck, Session.getScriptTimeZone(), "dd/MM/yyyy") : "DONE";
            let dataStr = data[i] && data[i][0] ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd/MM/yyyy") : "DONE";

            if (curentDateCheckStr == dataStr || dataStr == 'DONE') {
                coutSame += 1;
                if (i == data.length - 1 && coutSame != 1) {
                    let valueOld = sheet.getRange(startMergeRow, 15, coutSame, 1).getValue();
                    sheet.getRange(startMergeRow, 15, coutSame, 1).clearContent().merge().setValue(valueOld).setFontWeight("bold");
                    var data2 = sheet.getRange(startMergeRow, 17, coutSame, 1).getValues();
                    var sum2 = data2.reduce((acc, val) => acc + (val[0] || 0), 0);
                    sheet.getRange(startMergeRow, 18, coutSame, 1).clearContent().merge().setValue(sum2).setNumberFormat("#,##0");
                    if (currentColor == 1) {
                        hangColor = "#d9d2e9";
                        currentColor = 2;
                    } else {
                        hangColor = "#ffffff";
                        currentColor = 1;
                    }
                    sheet.getRange(startMergeRow, 15, coutSame, 4).setBackground(hangColor);
                    startMergeRow += coutSame;
                }
            } else {
                if (coutSame != 1) {
                    let valueOld = sheet.getRange(startMergeRow, 15, coutSame, 1).getValue();
                    sheet.getRange(startMergeRow, 15, coutSame, 1).clearContent().merge().setValue(valueOld).setFontWeight("bold");
                    var data1 = sheet.getRange(startMergeRow, 17, coutSame, 1).getValues();
                    var sum1 = data1.reduce((acc, val) => acc + (val[0] || 0), 0);
                    sheet.getRange(startMergeRow, 18, coutSame, 1).clearContent().merge().setValue(sum1).setNumberFormat("#,##0");
                    if (currentColor == 1) {
                        hangColor = "#d9d2e9";
                        currentColor = 2;
                    } else {
                        hangColor = "#ffffff";
                        currentColor = 1;
                    }
                    sheet.getRange(startMergeRow, 15, coutSame, 4).setBackground(hangColor);
                    startMergeRow += coutSame;
                } else {
                    if (currentColor == 1) {
                        hangColor = "#d9d2e9";
                        currentColor = 2;
                    } else {
                        hangColor = "#ffffff";
                        currentColor = 1;
                    }
                    sheet.getRange(startMergeRow, 15, coutSame, 4).setBackground(hangColor);
                    startMergeRow += 1;
                }
                coutSame = 1;
                curentDateCheck = data[i][0];
            }
        }
        // Gộp tất cả ô lại và set viền cùng lúc
        sheet.getRange(dongBatDauLuong, cotBatDauLuong, lastRow - dongBatDauLuong + 3, 5 + listThoChinh.length + listThoPhu.length).setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setVerticalAlignment("middle"); // (3, col): bắt đầu từ hàng 3, cột col, chiều cao là 2 hàng và 1 cột
        // end merge cột ngày giống nhau
        //  ui.alert("Tính lương đã xong.");

    }
}

function exportSheetToPdfAndEmail() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Chọn sheet hiện tại
    var sheetId = sheet.getSheetId(); // Lấy ID của sheet

    // Lấy ID của file Google Sheets
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

    // Thiết lập các tham số cho PDF
    var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
        'format=pdf&' + // Định dạng PDF
        'size=A4&' + // Kích thước giấy
        'portrait=true&' + // In dọc
        'fitw=true&' + // Phù hợp với chiều rộng
        'sheetnames=false&' + // Không hiển thị tên sheet
        'printtitle=false&' + // Không hiển thị tiêu đề
        'pagenumbers=true&' + // Hiển thị số trang
        'gridlines=false&' + // Ẩn đường lưới
        'fzr=false&' + // Không đông cột
        'gid=' + sheetId; // ID của sheet

    // Tạo request cho file PDF
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });

    // Tạo tệp PDF từ dữ liệu trả về
    var pdfBlob = response.getBlob().setName(sheet.getName() + ".pdf");

    // Gửi email với file PDF đính kèm
    var email = "mrthanh260801@gmail.com"; // Thay bằng email người nhận
    var subject = "Báo cáo PDF từ Google Sheets";
    var body = "Đây là báo cáo của bạn dưới dạng file PDF.";
    MailApp.sendEmail(email, subject, body, {
        attachments: [pdfBlob]
    });

    Logger.log("Email đã được gửi với file PDF đính kèm.");
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("🚀 Moon Hair")
        .addItem("Tạo bảng lương", "createOrUpdateSheetLuong")
        .addItem("Tính lương", "tinhLuong")
        .addItem("Cài tháng mặc định", "setThangDefault")
        .addItem("Tạo sheet mẫu", "createTables")
        .addToUi();
}

function filterMultipleColumns() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getFilter().remove(); // Xóa bộ lọc nếu có  
    var range = sheet.getRange("U4:S"); // Lấy toàn bộ dữ liệu
    var filter = range.createFilter(); // Tạo bộ lọc

    // Lọc cột B: chỉ hiển thị các số từ 100 đến 500
    filter.setColumnFilterCriteria(21, SpreadsheetApp.newFilterCriteria()
        .whenNumberBetween(0, 500000000000000000)
        .build());
}

// Hàm tạo file PDF từ vùng chọn
function createPDF() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Chọn sheet hiện tại
    var sheetId = sheet.getSheetId(); // Lấy ID của sheet

    // Lấy ID của file Google Sheets
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

    // Thiết lập các tham số cho PDF
    var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
        'format=pdf&' + // Định dạng PDF
        'size=A4&' + // Kích thước giấy
        'portrait=true&' + // In dọc
        'fitw=true&' + // Phù hợp với chiều rộng
        'sheetnames=false&' + // Không hiển thị tên sheet
        'printtitle=false&' + // Không hiển thị tiêu đề
        'pagenumbers=true&' + // Hiển thị số trang
        'gridlines=false&' + // Ẩn đường lưới
        'fzr=false&' + // Không đông cột
        'gid=' + sheetId; // ID của sheet

    // Tạo request cho file PDF
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });

    // Tạo tệp PDF từ dữ liệu trả về
    var blob = response.getBlob().setName(sheet.getName() + ".pdf");
    // var blob = sheetToPDF(sheet, range);
    var folder = DriveApp.getFolderById("1-28iaodg7mLLo0GgkOe4teSUEG-HmP5m"); // ID thư mục Drive để lưu
    var file = folder.createFile(blob).setName("Exported_PDF.pdf");

    return file.getUrl(); // Trả về URL của file PDF
}

// Hàm tạo file PDF từ vùng chọn
function createPDF() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange("A1:E10"); // Chọn vùng cần xuất PDF (có thể đổi)

    var folder = DriveApp.getFolderById("1-28iaodg7mLLo0GgkOe4teSUEG-HmP5m"); // ID thư mục Drive để lưu file
    var blob = sheetToPDF(sheet, range);
    var file = folder.createFile(blob).setName("Exported_PDF.pdf");

    return file.getUrl(); // Trả về URL của file PDF
}

// Chuyển Sheet thành PDF Blob
function sheetToPDF(sheet, range) {
    var spreadsheet = sheet.getParent();
    var sheetId = sheet.getSheetId();
    var rangeA1 = range.getA1Notation();

    var url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?`;
    url += `format=pdf&gid=${sheetId}&range=${rangeA1}`;
    url += "&portrait=true&size=A4";

    var params = {
        muteHttpExceptions: true,
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
    };
    var response = UrlFetchApp.fetch(url, params);

    return response.getBlob().setName("Exported_PDF.pdf");
}

function setThangDefault() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("sheetActive", sheet.getName());
}

function doGet(e) {
    var scriptProperties = PropertiesService.getScriptProperties();
    var value = scriptProperties.getProperty("sheetActive");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(value);
    let dongBatDau = 3;
    let cotBatDau = 15;
    let cotCuoiCung = sheet.getLastColumn() - cotBatDau;
    let dongCuoiCung = sheet.getLastRow() - 2;

    Logger.log(cotCuoiCung)
    Logger.log(dongCuoiCung)

    var data = sheet.getRange(dongBatDau, cotBatDau, dongCuoiCung, cotCuoiCung).getValues();
    Logger.log(value)
    Logger.log(JSON.stringify(data))

    var jsonOutput = ContentService.createTextOutput(JSON.stringify(data));
    jsonOutput.setMimeType(ContentService.MimeType.JSON);

    Logger.log(jsonOutput)
    return jsonOutput;

}

function createTables() {
    var ui = SpreadsheetApp.getUi(); // Lấy đối tượng UI
    // Hiển thị hộp thoại xác nhận
    var response = ui.alert("Xác nhận", "Có luốn tạo bảng mẫu?, lưu ý tạo sheet mới trước!", ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

        // Xóa nội dung cũ (nếu có)
        sheet.clear();

        // Tạo tiêu đề chính "Tháng 2/2025"
        sheet.getRange("A1").setValue("Tháng 2/2025").setFontWeight("bold").setFontSize(14);

        // Tiêu đề bảng chính
        var headers = ["Ngày", "Tên khách", "Tiền bill", "Phương thức", "Thợ chính", "Thợ phụ", "Ghi chú", "Tính lương"];
        sheet.getRange("A4:H4").setValues([headers]).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#DDDDDD");

        // Định dạng dropdown cho "Phương thức", "Thợ chính", "Thợ phụ"
        var validation = SpreadsheetApp.newDataValidation().requireValueInList(["Tiền mặt", "Chuyển khoản"]).build();
        sheet.getRange("D5:D20").setDataValidation(validation);

        var rangeListThoChinh = sheet.getRange("J4:J20");
        var validationThoChinh = SpreadsheetApp.newDataValidation().requireValueInRange(rangeListThoChinh).build();

        var rangeListThoPhu = sheet.getRange("L4:L20");
        var validationThoPhu = SpreadsheetApp.newDataValidation().requireValueInRange(rangeListThoPhu).build();

        sheet.getRange("E5:E20").setDataValidation(validationThoChinh);
        sheet.getRange("F5:F20").setDataValidation(validationThoPhu);

        // Tạo bảng Nhân viên tháng 11
        sheet.getRange("J1").setValue("Nhân viên tháng 11").setFontWeight("bold").setFontSize(12);
        sheet.getRange("J3:K3").setValues([["Thợ chính", "Lương (%)"]]).setFontWeight("bold").setBackground("#DDDDDD");
        sheet.getRange("L3:M3").setValues([["Thợ phụ", "Lương (%)"]]).setFontWeight("bold").setBackground("#DDDDDD");

        // Tô màu các ô ví dụ
        sheet.getRange("J4").setBackground("#FFD700"); // Vàng
        sheet.getRange("J6").setBackground("#00FF00"); // Xanh lá
        sheet.getRange("J7").setBackground("#FFFF00"); // Vàng nhạt
        sheet.getRange("L5").setBackground("#D3A4C2"); // Hồng nhạt

        // Căn giữa tiêu đề
        sheet.getRange("J4:M4").setHorizontalAlignment("center");

        sheet.getRange("A4:H20").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setVerticalAlignment("middle");
        sheet.getRange("J3:M20").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setVerticalAlignment("middle");
        sheet.getRange("A1:b1").merge().setHorizontalAlignment("left");

        sheet.getRange("A3:a4").merge();
        sheet.getRange("b3:b4").merge();
        sheet.getRange("c3:c4").merge();
        sheet.getRange("d3:d4").merge();
        sheet.getRange("e3:e4").merge();
        sheet.getRange("f3:f4").merge()
        sheet.getRange("g3:g4").merge()
        sheet.getRange("h3:h4").merge()
    }
}
