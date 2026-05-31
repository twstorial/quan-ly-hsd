// Tải dữ liệu khi mở trang
document.addEventListener("DOMContentLoaded", displayItems);

// 1. Thêm món hàng mới
function addItem() {
    const name = document.getElementById('itemName').value;
    const date = document.getElementById('expiryDate').value;

    if (!name || !date) {
        alert("Vui lòng nhập đủ tên và ngày!");
        return;
    }

    const items = JSON.parse(localStorage.getItem('hsd_items')) || [];
    items.push({
        id: Date.now() + Math.random(),
        name: name,
        date: date
    });

    localStorage.setItem('hsd_items', JSON.stringify(items));
    displayItems();
    
    document.getElementById('itemName').value = "";
    document.getElementById('expiryDate').value = "";
}

// 2. Hiển thị danh sách lên bảng
function displayItems() {
    const tableBody = document.getElementById('itemList');
    tableBody.innerHTML = "";
    
    let items = JSON.parse(localStorage.getItem('hsd_items')) || [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Lấy giá trị từ các ô chọn Sort
    const sortField = document.getElementById('sortField').value;
    const sortOrder = document.getElementById('sortOrder').value;

    // LOGIC SẮP XẾP
    items.sort((a, b) => {
        let valA, valB;

        if (sortField === 'name') {
            valA = a.name.toLowerCase();
            valB = b.name.toLowerCase();
        } else if (sortField === 'date' || sortField === 'diff') {
            // Cả ngày hết hạn và số ngày còn lại đều dựa trên giá trị thời gian
            valA = new Date(a.date).getTime();
            valB = new Date(b.date).getTime();
        }

        if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
        if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
        return 0;
    });

    // VẼ BẢNG SAU KHI ĐÃ SORT
    items.forEach(item => {
        const expDate = new Date(item.date);
        const diffTime = expDate - today;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        // Định dạng dd/mm/yy để hiển thị
        const day = String(expDate.getDate()).padStart(2, '0');
        const month = String(expDate.getMonth() + 1).padStart(2, '0');
        const year = String(expDate.getFullYear()).slice(-2);
        const formattedDate = `${day}/${month}/${year}`;

        let statusText, statusClass, countdownDisplay;
        if (diffDays < 0) {
            statusText = "Đã hết hạn";
            statusClass = "status-expired";
            countdownDisplay = `<span class="expired-text">Quá ${Math.abs(diffDays)} ngày</span>`;
        } else if (diffDays === 0) {
            statusText = "Hết hạn hôm nay";
            statusClass = "status-warning";
            countdownDisplay = `<span class="countdown-text">Hôm nay</span>`;
        } else {
            statusText = diffDays <= 3 ? "Sắp hết hạn" : "Còn hạn";
            statusClass = diffDays <= 3 ? "status-warning" : "status-ok";
            countdownDisplay = `<span class="countdown-text">${diffDays} ngày nữa</span>`;
        }

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.name}</td>
            <td>${formattedDate}</td>
            <td>${countdownDisplay}</td>
            <td><span class="status ${statusClass}">${statusText}</span></td>
            <td><button class="delete-btn" onclick="deleteItem(${item.id})">Xóa</button></td>
        `;
        tableBody.appendChild(row);
    });
}

// 3. Xóa món hàng
function deleteItem(id) {
    let items = JSON.parse(localStorage.getItem('hsd_items')) || [];
    items = items.filter(item => item.id !== id);
    localStorage.setItem('hsd_items', JSON.stringify(items));
    displayItems();
}

// 4. XUẤT FILE EXCEL (Lưu vào folder máy)
function exportToExcel() {
    const items = JSON.parse(localStorage.getItem('hsd_items')) || [];
    if (items.length === 0) return alert("Không có dữ liệu!");

    const dataForExcel = items.map(item => ({
        "Tên Hàng": item.name,
        "Ngày Hết Hạn": item.date
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataForExcel);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "DanhSachHSD");
    
    // Trình duyệt sẽ yêu cầu chọn nơi lưu file
    XLSX.writeFile(workbook, "QuanLyHanSuDung.xlsx");
}

// 5. NHẬP FILE EXCEL (Đọc từ máy tính lên web)
function importFromExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        const items = JSON.parse(localStorage.getItem('hsd_items')) || [];

        jsonData.forEach(row => {
            if (row["Tên Hàng"] && row["Ngày Hết Hạn"]) {
                items.push({
                    id: Date.now() + Math.random(),
                    name: row["Tên Hàng"],
                    date: row["Ngày Hết Hạn"]
                });
            }
        });

        localStorage.setItem('hsd_items', JSON.stringify(items));
        displayItems();
        alert("Đã tải dữ liệu thành công!");
        event.target.value = ''; // Reset input
    };
    reader.readAsArrayBuffer(file);
}