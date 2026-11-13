
/* ============================================================
   FILE: SCRIPT.JS
   WEBSITE: TRA CỨU THÔNG SỐ CÀI ĐẶT RƠ-LE HPDQ1

   CHỨC NĂNG CHÍNH:
   1. Đọc file Excel (.xlsx).
   2. Lấy danh sách các Sheet trong Excel.
   3. Hiển thị dữ liệu Excel lên bảng HTML.
   4. Chuyển đổi giữa các Sheet.
   5. Tìm kiếm thông tin trong bảng.
   6. Cập nhật số lượng dữ liệu đang hiển thị.
   7. Xử lý thông báo lỗi.

   THƯ VIỆN SỬ DỤNG: SheetJS (XLSX)
   ============================================================ */


/* ============================================================
   PHẦN 1: KẾT NỐI JAVASCRIPT VỚI CÁC THÀNH PHẦN HTML

   document.getElementById():
   Tìm phần tử HTML thông qua ID được khai báo trong index.html.
   ============================================================ */

const fileInput = document.getElementById("excelFile");

const sheetSelect = document.getElementById("sheetSelect");

const searchInput = document.getElementById("searchInput");

const tableContainer = document.getElementById("tableContainer");

const fileName = document.getElementById("fileName");

const statusBadge = document.getElementById("statusBadge");

const currentSheet = document.getElementById("currentSheet");

const visibleRows = document.getElementById("visibleRows");

const sheetCount = document.getElementById("sheetCount");

const tableTitle = document.getElementById("tableTitle");

const tableHint = document.getElementById("tableHint");

const clearSearchButton = document.getElementById("clearSearch");

const resetViewButton = document.getElementById("resetView");


/* ============================================================
   PHẦN 2: KHAI BÁO BIẾN LƯU TRỮ DỮ LIỆU

   workbook:
   Lưu toàn bộ nội dung file Excel sau khi được đọc.

   Ban đầu workbook = null vì người dùng chưa chọn file.
   ============================================================ */

let workbook = null;


/* ============================================================
   PHẦN 3: ĐỌC FILE EXCEL

   Sự kiện change:
   Được kích hoạt khi người dùng chọn file Excel.

   FileReader:
   Đọc nội dung file Excel từ máy tính vào trình duyệt.
   ============================================================ */

fileInput.addEventListener("change", function (event) {

    const file = event.target.files[0];

    // Kiểm tra người dùng đã chọn file hay chưa
    if (!file) return;


    // Kiểm tra thư viện SheetJS đã được tải chưa
    if (typeof XLSX === "undefined") {

        showError(
            "Không tải được thư viện SheetJS. Hãy kiểm tra kết nối Internet."
        );

        return;
    }


    // Hiển thị tên file Excel đang được chọn
    fileName.textContent = file.name;

    statusBadge.textContent = "Đang đọc file...";


    // Khởi tạo đối tượng đọc file
    const reader = new FileReader();


    /* --------------------------------------------------------
       PHẦN 3.1: XỬ LÝ SAU KHI ĐỌC FILE THÀNH CÔNG
       -------------------------------------------------------- */

    reader.onload = function (e) {

        try {

            // Đọc dữ liệu Excel bằng thư viện SheetJS
            workbook = XLSX.read(e.target.result, {

                type: "array",

                cellStyles: true,

                cellDates: true

            });


            // Kiểm tra workbook có Sheet hay không
            if (!workbook.SheetNames.length) {

                showError("File Excel không có Sheet dữ liệu.");

                return;
            }


            /* ------------------------------------------------
               PHẦN 3.2: LẤY DANH SÁCH SHEET TỪ EXCEL

               SheetNames:
               Chứa tên tất cả các Sheet trong file Excel.

               Ví dụ:
               Bãi Liệu
               TKVV
               Luyện Cốc
               LT
               LG1
               CTD1
               CTD2
               CTD3
               QSP
               Sheet1
               ------------------------------------------------ */

            sheetSelect.innerHTML = "";


            workbook.SheetNames.forEach(function (name) {

                const option = document.createElement("option");

                option.value = name;

                option.textContent = name;

                sheetSelect.appendChild(option);

            });


            // Mở khóa các chức năng trên giao diện
            sheetSelect.disabled = false;

            searchInput.disabled = false;

            clearSearchButton.disabled = false;

            resetViewButton.disabled = false;


            // Cập nhật tổng số Sheet
            sheetCount.textContent = workbook.SheetNames.length;


            // Cập nhật trạng thái
            statusBadge.textContent = "Đã nạp dữ liệu";


            // Tự động hiển thị Sheet đầu tiên
            renderSheet(workbook.SheetNames[0]);


        } catch (error) {

            console.error(error);

            showError(
                "Không đọc được file Excel. Hãy kiểm tra định dạng file."
            );

        }

    };


    /* --------------------------------------------------------
       PHẦN 3.3: XỬ LÝ KHI KHÔNG ĐỌC ĐƯỢC FILE
       -------------------------------------------------------- */

    reader.onerror = function () {

        showError("Trình duyệt không thể đọc file Excel.");

    };


    // Bắt đầu đọc file Excel
    reader.readAsArrayBuffer(file);

});


/* ============================================================
   PHẦN 4: CHUYỂN ĐỔI GIỮA CÁC SHEET

   Khi người dùng chọn Sheet khác trên giao diện,
   JavaScript sẽ gọi hàm renderSheet() để hiển thị dữ liệu.

   Ví dụ:
   Chọn Bãi Liệu -> Hiển thị dữ liệu Bãi Liệu.
   Chọn TKVV -> Hiển thị dữ liệu TKVV.
   ============================================================ */

sheetSelect.addEventListener("change", function () {

    if (workbook && this.value) {

        // Xóa từ khóa tìm kiếm cũ
        searchInput.value = "";

        // Hiển thị Sheet được chọn
        renderSheet(this.value);

    }

});


/* ============================================================
   PHẦN 5: HIỂN THỊ DỮ LIỆU EXCEL LÊN HTML

   Hàm renderSheet(sheetName):

   Đầu vào:
   Tên Sheet Excel cần hiển thị.

   Xử lý:
   1. Lấy dữ liệu Sheet tương ứng.
   2. Chuyển dữ liệu Excel thành bảng HTML.
   3. Đưa bảng HTML vào tableContainer.
   4. Cập nhật tên Sheet và thống kê.

   sheet_to_html():
   Hàm của thư viện SheetJS dùng để chuyển Sheet thành HTML.
   ============================================================ */

function renderSheet(sheetName) {

    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet) return;


    // Chuyển dữ liệu Excel thành mã HTML
    const html = XLSX.utils.sheet_to_html(worksheet, {

        id: "excelDataTable",

        editable: false

    });


    // Hiển thị bảng HTML lên giao diện website
    tableContainer.innerHTML = html;


    // Cập nhật tên Sheet đang mở
    currentSheet.textContent = sheetName;

    tableTitle.textContent = sheetName;


    // Cập nhật hướng dẫn hiển thị
    tableHint.textContent =
        "Dữ liệu được chuyển trực tiếp từ Sheet Excel. Cuộn ngang / dọc để xem toàn bộ bảng.";


    // Cập nhật số hàng hiển thị
    applySearch();

}


/* ============================================================
   PHẦN 6: CHỨC NĂNG TÌM KIẾM DỮ LIỆU

   Người dùng nhập từ khóa vào ô tìm kiếm.

   JavaScript sẽ:
   1. Lấy nội dung từ khóa.
   2. Đọc nội dung từng hàng trong bảng.
   3. So sánh với từ khóa.
   4. Ẩn các hàng không phù hợp.

   Ví dụ:
   Nhập "7SJ62" -> Hiển thị các hàng có chứa 7SJ62.
   ============================================================ */

searchInput.addEventListener("input", applySearch);


function applySearch() {

    const table = tableContainer.querySelector("table");

    if (!table) return;


    // Lấy từ khóa tìm kiếm
    const keyword = searchInput.value
        .trim()
        .toLocaleLowerCase("vi");


    // Lấy toàn bộ hàng trong bảng
    const rows = Array.from(table.querySelectorAll("tr"));


    let shown = 0;


    // Kiểm tra từng hàng
    rows.forEach(function (row) {

        const rowText = row.textContent
            .toLocaleLowerCase("vi");


        // Nếu không có từ khóa hoặc hàng chứa từ khóa
        const matched = !keyword || rowText.includes(keyword);


        // Hiển thị hoặc ẩn hàng
        row.hidden = !matched;


        if (matched) {

            shown++;

        }

    });


    // Cập nhật số hàng đang hiển thị
    visibleRows.textContent = shown + " / " + rows.length;


    // Cập nhật thông báo tìm kiếm
    if (keyword) {

        tableHint.textContent =
            "Đang lọc theo từ khóa: " + searchInput.value;

    } else {

        tableHint.textContent =
            "Đang hiển thị toàn bộ dữ liệu của Sheet.";

    }

}


/* ============================================================
   PHẦN 7: NÚT XÓA TỪ KHÓA TÌM KIẾM

   Khi nhấn nút "Xóa tìm kiếm":
   - Xóa nội dung ô tìm kiếm.
   - Hiển thị lại toàn bộ dữ liệu.
   ============================================================ */

clearSearchButton.addEventListener("click", function () {

    searchInput.value = "";

    applySearch();

    searchInput.focus();

});


/* ============================================================
   PHẦN 8: NÚT HIỂN THỊ TẤT CẢ HÀNG

   Chức năng tương tự nút xóa tìm kiếm.
   Giúp người dùng quay lại trạng thái xem toàn bộ bảng.
   ============================================================ */

resetViewButton.addEventListener("click", function () {

    searchInput.value = "";

    applySearch();

});


/* ============================================================
   PHẦN 9: XỬ LÝ THÔNG BÁO LỖI

   Khi xảy ra lỗi:
   - Cập nhật trạng thái website.
   - Hiển thị nội dung lỗi.
   - Đặt lại thông tin thống kê.
   ============================================================ */

function showError(message) {

    statusBadge.textContent = "Có lỗi";

    tableContainer.textContent = message;

    currentSheet.textContent = "—";

    visibleRows.textContent = "—";

}
