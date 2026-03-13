package controller;

import util.ExcelProcessor;

import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.MultipartConfig;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.Part;

import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Servlet that accepts an Excel file upload, extracts the customs-declaration
 * fields "Mã số hàng hóa" and "Mô tả hàng hóa" from every sheet, writes the
 * collected data into a new "Result" sheet, and returns the modified workbook
 * as a download.
 */
@WebServlet("/excel")
@MultipartConfig(maxFileSize = 10 * 1024 * 1024)   // 10 MB limit
public class ExcelServlet extends HttpServlet {

    /** Shows the upload form. */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        request.getRequestDispatcher("excel.jsp").forward(request, response);
    }

    /** Processes the uploaded Excel file and streams the result back. */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        Part filePart = request.getPart("excelFile");
        if (filePart == null || filePart.getSize() == 0) {
            request.setAttribute("error", "Vui lòng chọn tệp Excel trước khi tải lên.");
            request.getRequestDispatcher("excel.jsp").forward(request, response);
            return;
        }

        String fileName = getFileName(filePart);
        if (!isExcelFile(fileName)) {
            request.setAttribute("error", "Tệp không hợp lệ. Vui lòng tải lên tệp .xlsx hoặc .xls.");
            request.getRequestDispatcher("excel.jsp").forward(request, response);
            return;
        }

        // Validate actual file content using Apache POI magic bytes
        try (InputStream probe = new BufferedInputStream(filePart.getInputStream())) {
            FileMagic magic = FileMagic.valueOf(probe);
            if (magic != FileMagic.OOXML && magic != FileMagic.OLE2) {
                request.setAttribute("error",
                        "Nội dung tệp không hợp lệ. Vui lòng tải lên tệp Excel thực sự (.xlsx hoặc .xls).");
                request.getRequestDispatcher("excel.jsp").forward(request, response);
                return;
            }
        }

        // Determine output MIME type and file extension
        String outputMime = fileName.toLowerCase().endsWith(".xls")
                ? "application/vnd.ms-excel"
                : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        String outputExt = fileName.toLowerCase().endsWith(".xls") ? ".xls" : ".xlsx";
        String downloadName = "result" + outputExt;

        response.setContentType(outputMime);
        response.setHeader("Content-Disposition", "attachment; filename=\"" + downloadName + "\"");
        response.setCharacterEncoding("UTF-8");

        try (InputStream is = filePart.getInputStream()) {
            ExcelProcessor.process(is, response.getOutputStream(), fileName);
        } catch (Exception e) {
            // Reset the response if headers haven't been committed yet
            if (!response.isCommitted()) {
                response.reset();
                request.setAttribute("error", "Đã xảy ra lỗi khi xử lý tệp: " + e.getMessage());
                request.getRequestDispatcher("excel.jsp").forward(request, response);
            }
        }
    }

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    /** Extracts the original file name from the multipart {@link Part}. */
    private String getFileName(Part part) {
        String contentDisposition = part.getHeader("content-disposition");
        if (contentDisposition == null) {
            return "";
        }
        for (String token : contentDisposition.split(";")) {
            token = token.trim();
            if (token.startsWith("filename")) {
                return token.substring(token.indexOf('=') + 1).trim().replace("\"", "");
            }
        }
        return "";
    }

    private boolean isExcelFile(String fileName) {
        if (fileName == null || fileName.isEmpty()) {
            return false;
        }
        String lower = fileName.toLowerCase();
        return lower.endsWith(".xlsx") || lower.endsWith(".xls");
    }
}
