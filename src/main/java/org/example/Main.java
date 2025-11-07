package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String fixedDate = "20251105";

        String url = "jdbc:sqlserver://MAYCHU:1433;databaseName=ITDVAPCF;encrypt=true;trustServerCertificate=true;";
        String user = "sa";
        String password = "1q2w3e4r5t!@#$%aA@th";

        String sql = """
            WITH Giaodich_Filtered AS (
                SELECT DISTINCT object_id, ten_loai_giao_dich
                FROM vwGiao_Dich
                WHERE 
                    Convert(VARCHAR(10), Ngay, 112) = ?
                    AND ten_loai_giao_dich IN (N'Xuất tài sản thế chấp', N'Nhập tài sản thế chấp')
                    AND object_id IS NOT NULL
            )
            SELECT 
                TSTC.ChuTS_Hoten AS [Họ và tên],
                TSTC.ChuTS_Diachi AS [Địa chỉ],
                TSTC.tstc_soluong AS [Số lượng],
                TSTC.tstc_ten AS [Seri],
                GD.ten_loai_giao_dich AS [Loại giao dịch]
            FROM Tdung_Taisanthechap TSTC
            INNER JOIN Giaodich_Filtered GD ON GD.object_id = TSTC.TSTC_ID
            """;

        List<TaiSanTheChap> danhSach = new ArrayList<>();

        try (Connection conn = DriverManager.getConnection(url, user, password);
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, fixedDate);

            try (ResultSet rs = stmt.executeQuery()) {
                ResultSetMetaData meta = rs.getMetaData();
                int columnCount = meta.getColumnCount();
                int rowCount = 0;

                System.out.println("📋 Báo cáo tài sản thế chấp:");
                while (rs.next()) {
                    rowCount++;
                    TaiSanTheChap item = new TaiSanTheChap();
                    item.hoTen = rs.getString("Họ và tên");
                    item.diaChi = rs.getString("Địa chỉ");
                    item.soLuong = rs.getInt("Số lượng");
                    item.seri = rs.getString("Seri");
                    item.loaiGiaoDich = rs.getString("Loại giao dịch");
                    danhSach.add(item);

                    StringBuilder record = new StringBuilder("🔹 ");
                    for (int i = 1; i <= columnCount; i++) {
                        record.append(meta.getColumnLabel(i)).append(": ").append(rs.getString(i)).append(" | ");
                    }
                    System.out.println(record.toString());
                }

                System.out.println("📊 Tổng số bản ghi báo cáo: " + rowCount);
                if (rowCount == 0) {
                    System.out.println("⚠️ Không có dữ liệu phù hợp với ngày và loại giao dịch đã chọn.");
                    return;
                }
            }
        } catch (SQLException e) {
            System.out.println("❌ Lỗi kết nối hoặc thực thi truy vấn:");
            e.printStackTrace();
            return;
        }

        // Ghi dữ liệu vào Excel
        try (FileInputStream fis = new FileInputStream("template/Lệnh nhập xuất kho.xlsx");
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheetXuat = workbook.getSheet("Xuất TSTC");
            Sheet sheetNhap = workbook.getSheet("Nhập TSTC");

            int startRowXuat = 24;
            int startRowNhap = 24;
            int sttXuat = 1, sttNhap = 1;

            for (TaiSanTheChap item : danhSach) {
                boolean isXuat = item.loaiGiaoDich != null && item.loaiGiaoDich.contains("Xuất");
                Sheet targetSheet = isXuat ? sheetXuat : sheetNhap;
                int rowIndex = isXuat ? startRowXuat++ : startRowNhap++;
                Row row = targetSheet.createRow(rowIndex);

                row.createCell(0).setCellValue(isXuat ? sttXuat++ : sttNhap++);
                row.createCell(1).setCellValue(item.hoTen);
                row.createCell(2).setCellValue(item.diaChi);
                row.createCell(3).setCellValue(item.soLuong);
                row.createCell(4).setCellValue(item.seri);
            }

            // Tạo thư mục theo ngày
            String folderName = "output/" + fixedDate;
            File folder = new File(folderName);
            if (!folder.exists()) folder.mkdirs();

            // Ghi file ra thư mục
            String outputFile = folderName + "/Lệnh nhập xuất kho_" + fixedDate + ".xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
                System.out.println("✅ Đã xuất file Excel: " + outputFile);
            }

        } catch (IOException e) {
            System.out.println("❌ Lỗi xử lý file Excel:");
            e.printStackTrace();
        }
    }
}
