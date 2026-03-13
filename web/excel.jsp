<%-- 
    Document   : excel
    Description: Upload an Excel workbook to extract customs-declaration fields
                 "Mã số hàng hóa" and "Mô tả hàng hóa" from every sheet.
--%>

<%@page contentType="text/html" pageEncoding="UTF-8"%>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Trích xuất dữ liệu tờ khai hải quan</title>
        <style>
            body {
                margin: 60px auto;
                max-width: 700px;
                font-family: Arial, sans-serif;
                background-color: #f4f4f4;
                color: #333;
            }

            h1 {
                text-align: center;
                color: #007bff;
            }

            .card {
                background: #fff;
                border-radius: 6px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.12);
                padding: 32px 40px;
                margin-top: 24px;
            }

            label {
                font-weight: bold;
                display: block;
                margin-bottom: 8px;
            }

            input[type="file"] {
                width: 100%;
                padding: 8px;
                margin-bottom: 16px;
                box-sizing: border-box;
                border: 1px solid #ccc;
                border-radius: 4px;
            }

            input[type="submit"] {
                width: 100%;
                padding: 10px;
                background-color: #007bff;
                color: #fff;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 16px;
            }

            input[type="submit"]:hover {
                background-color: #0056b3;
            }

            .error {
                color: red;
                margin-bottom: 16px;
                font-weight: bold;
            }

            .info {
                font-size: 13px;
                color: #666;
                margin-top: 16px;
            }

            ul.info-list {
                margin: 6px 0 0 18px;
                font-size: 13px;
                color: #555;
            }
        </style>
    </head>
    <body>
        <h1>Trích xuất dữ liệu tờ khai hải quan</h1>

        <div class="card">
            <% String error = (String) request.getAttribute("error"); %>
            <% if (error != null && !error.isEmpty()) { %>
                <p class="error"><%= error %></p>
            <% } %>

            <form action="excel" method="post" enctype="multipart/form-data">
                <label for="excelFile">Chọn tệp Excel (.xlsx hoặc .xls):</label>
                <input type="file" id="excelFile" name="excelFile" accept=".xlsx,.xls" required />
                <input type="submit" value="Tải lên &amp; Xử lý" />
            </form>

            <p class="info">
                <strong>Cách hoạt động:</strong>
            </p>
            <ul class="info-list">
                <li>Hệ thống sẽ quét tất cả các sheet trong workbook.</li>
                <li>Tìm kiếm nhãn <strong>"Mã số hàng hóa"</strong> và <strong>"Mô tả hàng hóa"</strong>.</li>
                <li>Trích xuất giá trị trong ô kế bên các nhãn đó.</li>
                <li>Tổng hợp kết quả vào sheet <strong>"Result"</strong> và trả về tệp Excel mới.</li>
            </ul>
        </div>
    </body>
</html>
