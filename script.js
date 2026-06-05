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

// // 3. Xóa món hàng
// function deleteItem(id) {
//     let items = JSON.parse(localStorage.getItem('hsd_items')) || [];
//     items = items.filter(item => item.id !== id);
//     localStorage.setItem('hsd_items', JSON.stringify(items));
//     displayItems();
// }

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

        if (jsonData.length === 0) {
            alert("File không có dữ liệu!");
            return;
        }

        // BƯỚC THAY ĐỔI: Tạo một mảng trống mới hoàn toàn (Xóa sạch bộ nhớ cũ)
        const newItems = [];

        jsonData.forEach(row => {
            if (row["Tên Hàng"] && row["Ngày Hết Hạn"]) {
                newItems.push({
                    id: Date.now() + Math.random(), // Tạo ID mới cho các món
                    name: row["Tên Hàng"],
                    date: row["Ngày Hết Hạn"]
                });
            }
        });

        // Ghi đè mảng mới này vào localStorage
        localStorage.setItem('hsd_items', JSON.stringify(newItems));
        
        displayItems(); // Vẽ lại bảng
        
        // Nếu đang mở block tìm kiếm thì cập nhật hoặc đóng nó đi
        document.getElementById('searchResultBlock').style.display = "none";
        document.getElementById('searchInput').value = "";

        alert("Tải file thành công! Dữ liệu cũ đã được thay thế.");
        event.target.value = ''; 
    };
    reader.readAsArrayBuffer(file);
}

// 6. Hàm tìm kiếm món hàng
function searchItem() {
    const keyword = document.getElementById('searchInput').value.trim().toLowerCase();
    const resultBlock = document.getElementById('searchResultBlock');
    const resultContent = document.getElementById('searchResultContent');
    
    if (!keyword) {
        alert("Vui lòng nhập từ khóa tìm kiếm!");
        return;
    }

    const items = JSON.parse(localStorage.getItem('hsd_items')) || [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Lọc các món hàng có tên chứa từ khóa
    const filteredItems = items.filter(item => item.name.toLowerCase().includes(keyword));

    if (filteredItems.length === 0) {
        resultContent.innerHTML = "<p>Không tìm thấy món hàng nào phù hợp.</p>";
    } else {
        resultContent.innerHTML = ""; // Xóa kết quả cũ
        filteredItems.forEach(item => {
            const expDate = new Date(item.date);
            const diffDays = Math.ceil((expDate - today) / (1000 * 60 * 60 * 24));
            
            // Định dạng dd/mm/yy
            const day = String(expDate.getDate()).padStart(2, '0');
            const month = String(expDate.getMonth() + 1).padStart(2, '0');
            const year = String(expDate.getFullYear()).slice(-2);
            const formattedDate = `${day}/${month}/${year}`;

            const itemDiv = document.createElement('div');
            itemDiv.className = 'result-item';
            itemDiv.innerHTML = `
                <div>
                    <strong>${item.name}</strong> - <small>HSD: ${formattedDate}</small>
                </div>
                <div>
                    <span class="countdown-text">${diffDays < 0 ? 'Hết hạn' : diffDays + ' ngày nữa'}</span>
                </div>
            `;
            resultContent.appendChild(itemDiv);
        });
    }

    resultBlock.style.display = "block"; // Hiển thị block kết quả
}

// 6.2. Hàm làm mới (ẩn block tìm kiếm)
function clearSearch() {
    document.getElementById('searchInput').value = "";
    document.getElementById('searchResultBlock').style.display = "none";
}

// Lưu ý: Cần chỉnh sửa lại hàm deleteItem một chút 
// để nếu đang mở tìm kiếm mà xóa thì nó cập nhật cả hai.
function deleteItem(id) {
    // 1. Lấy thông tin món hàng dựa trên ID để hiển thị trong thông báo cho rõ ràng
    const items = JSON.parse(localStorage.getItem('hsd_items')) || [];
    const itemToDelete = items.find(item => item.id === id);
    const itemName = itemToDelete ? itemToDelete.name : "món hàng này";

    // 2. Hiển thị hộp thoại xác nhận (Confirm)
    const isConfirmed = confirm(`Bạn có chắc chắn muốn xóa món hàng "${itemName}" không?`);

    // 3. Nếu người dùng đồng ý (bấm OK) thì tiến hành xóa
    if (isConfirmed) {
        const updatedItems = items.filter(item => item.id !== id);
        localStorage.setItem('hsd_items', JSON.stringify(updatedItems));
        
        displayItems(); // Cập nhật lại bảng chính
        
        // Nếu block tìm kiếm đang mở, thực hiện lại tìm kiếm để cập nhật giao diện tìm kiếm
        if (document.getElementById('searchResultBlock').style.display === "block") {
            // Kiểm tra xem từ khóa hiện tại còn khớp món nào không
            const keyword = document.getElementById('searchInput').value.trim();
            if (keyword) {
                searchItem();
            }
        }
    }
    // Nếu người dùng bấm Cancel (Hủy), hàm sẽ dừng tại đây và không xóa gì cả.
}

// 7. Hàm xóa toàn bộ danh sách (Clear All)
function clearAllItems() {
    // 1. Kiểm tra xem hiện tại có dữ liệu để xóa hay không
    const items = JSON.parse(localStorage.getItem('hsd_items')) || [];
    if (items.length === 0) {
        alert("Danh sách hiện tại đã trống sẵn rồi!");
        return;
    }

    // 2. Hiển thị hộp thoại xác nhận (Confirm) với nội dung cảnh báo mạnh hơn
    const isConfirmed = confirm("⚠️ CẢNH BÁO: Bạn có chắc chắn muốn XÓA TOÀN BỘ danh sách món hàng không? \nHành động này không thể hoàn tác trừ khi bạn đã lưu file Excel!");

    // 3. Nếu đồng ý, tiến hành xóa sạch bộ nhớ và giao diện
    if (isConfirmed) {
        // Xóa hoàn toàn key dữ liệu trong localStorage
        localStorage.removeItem('hsd_items');
        
        // Vẽ lại bảng trống
        displayItems();
        
        // Đóng và xóa thanh tìm kiếm nếu đang mở
        document.getElementById('searchResultBlock').style.display = "none";
        document.getElementById('searchInput').value = "";
        
        alert("Đã xóa sạch toàn bộ danh sách thành công. Bạn có thể tạo danh sách mới!");
    }
}