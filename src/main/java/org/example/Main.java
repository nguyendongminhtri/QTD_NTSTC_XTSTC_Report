package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.sql.*;
import java.text.Normalizer;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    static class TaiSanTheChap {
        String hoTen;
        String diaChi;
        int soLuong;
        String seri;
        String loaiGiaoDich;
    }

    public static void main(String[] args) {
        String url = "jdbc:sqlserver://MAYCHU:1433;databaseName=ITDVAPCF;encrypt=true;trustServerCertificate=true;";
        String user = "sa";
        String password = "1q2w3e4r5t!@#$%aA@th";

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        LocalDate startDate = LocalDate.of(2025, 1, 1);
        LocalDate endDate = LocalDate.now();

        String sharedPath = "D:\\Bao Cao TSTC From 2025";
        List<String> danhSachFileDaXuat = new ArrayList<>();

        try (Connection conn = DriverManager.getConnection(url, user, password)) {
            System.out.println("✅ Kết nối DB thành công.");

            LocalDate cur = startDate;
            while (!cur.isAfter(endDate)) {
                String fixedDate = cur.format(formatter);
                System.out.println("🔄 Đang xử lý ngày: " + fixedDate);
                processDate(fixedDate, conn, danhSachFileDaXuat); // truyền kết nối vào
                cur = cur.plusDays(1);
            }

            System.out.println("🔁 Bắt đầu sao chép tất cả file đã xuất sang: " + sharedPath);
            for (String filePath : danhSachFileDaXuat) {
                File sourceFile = new File(filePath);
                int index = sourceFile.getAbsolutePath().indexOf("output");
                String relativePath = sourceFile.getAbsolutePath().substring(index + "output".length());
                File destFile = new File(sharedPath + File.separator + relativePath);
                destFile.getParentFile().mkdirs();

                try {
                    Files.copy(sourceFile.toPath(), destFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    System.out.println("📁 Đã sao chép: " + sourceFile.getName() + " → " + destFile.getAbsolutePath());
                } catch (IOException e) {
                    System.out.println("❌ Lỗi sao chép file: " + sourceFile.getName());
                    e.printStackTrace();
                }
            }

            System.out.println("✅ Hoàn tất sao chép các file.");
        } catch (SQLException e) {
            System.out.println("❌ Không thể kết nối DB:");
            e.printStackTrace();
        }
    }

    private static void processDate(String fixedDate, Connection conn, List<String> danhSachFileDaXuat) {
        String sql = """
        WITH Giaodich_Filtered AS (
            SELECT DISTINCT object_id, ten_loai_giao_dich
            FROM vwGiao_Dich
            WHERE 
                 CAST(Ngay AS DATE) = ?
                AND ten_loai_giao_dich IN (N'Xuất tài sản thế chấp', N'Nhập tài sản thế chấp', N'Xuất TS giữ hộ')
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

        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            stmt.setString(1, fixedDate);

            try (ResultSet rs = stmt.executeQuery()) {
                while (rs.next()) {
                    TaiSanTheChap item = new TaiSanTheChap();
                    item.hoTen = rs.getString("Họ và tên");
                    item.diaChi = rs.getString("Địa chỉ");
                    item.soLuong = rs.getInt("Số lượng");
                    item.seri = rs.getString("Seri");
                    item.loaiGiaoDich = rs.getString("Loại giao dịch");
                    danhSach.add(item);
                }
            }
        } catch (SQLException e) {
            System.out.println("❌ Lỗi truy vấn SQL cho ngày " + fixedDate + ":");
            e.printStackTrace();
            return;
        }

        System.out.println("📊 Số bản ghi ngày " + fixedDate + ": " + danhSach.size());

        try (InputStream is = Main.class.getClassLoader().getResourceAsStream("template/Lenh_Xuat_Nhap_TSTC.xlsx")) {
            if (is == null) {
                System.out.println("❌ Không tìm thấy file template trong JAR.");
                return;
            }

            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheetXuat = null;
            Sheet sheetNhap = null;

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                String name = workbook.getSheetName(i);
                String nameKhongDau = boDau(name).toLowerCase();

                if (nameKhongDau.contains("xuat")) {
                    sheetXuat = workbook.getSheetAt(i);
                } else if (nameKhongDau.contains("nhap")) {
                    sheetNhap = workbook.getSheetAt(i);
                }
            }

            List<TaiSanTheChap> danhSachXuat = new ArrayList<>();
            List<TaiSanTheChap> danhSachNhap = new ArrayList<>();

            for (TaiSanTheChap item : danhSach) {
                if (item.loaiGiaoDich != null && item.loaiGiaoDich.contains("Xuất")) {
                    danhSachXuat.add(item);
                } else {
                    danhSachNhap.add(item);
                }
            }

            if (sheetXuat != null) {
                ghiDuLieu(workbook, sheetXuat, danhSachXuat, 43);
            } else {
                System.out.println("⚠️ Không tìm thấy sheet chứa 'Xuất' trong tên.");
            }

            if (sheetNhap != null) {
                ghiDuLieu(workbook, sheetNhap, danhSachNhap, 43);
            } else {
                System.out.println("⚠️ Không tìm thấy sheet chứa 'Nhập' trong tên.");
            }

            String jarDirPath = new File(Main.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getParent();
            String year = fixedDate.substring(0, 4);
            String month = fixedDate.substring(4, 6);
            File outputFolder = new File(jarDirPath, "output/" + year + "/" + month + "/" + fixedDate);
            if (!outputFolder.exists() && !outputFolder.mkdirs()) {
                System.out.println("❌ Không thể tạo thư mục đầu ra: " + outputFolder.getAbsolutePath());
                return;
            }

            File outputFile = new File(outputFolder, "Nhap Xuat TSTC_" + fixedDate + ".xlsx");
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
                System.out.println("✅ Đã xuất file Excel: " + outputFile.getAbsolutePath());
                danhSachFileDaXuat.add(outputFile.getAbsolutePath());
            }
        } catch (Exception e) {
            System.out.println("❌ Lỗi xử lý file Excel cho ngày " + fixedDate + ":");
            e.printStackTrace();
        }
    }



    private static void ghiDuLieu(Workbook workbook, Sheet sheet, List<TaiSanTheChap> danhSach, int startRow) {
        if (!danhSach.isEmpty()) {
            int rowsToInsert = danhSach.size();
            int lastRow = sheet.getLastRowNum();
            if (lastRow >= startRow) {
                sheet.shiftRows(startRow, lastRow, rowsToInsert);
            }
        }

        // 👉 Font chữ 14pt
        Font normalFont = workbook.createFont();
        normalFont.setFontName("Times New Roman");
        normalFont.setFontHeightInPoints((short) 14);
        normalFont.setBold(false);

        CellStyle borderedStyle = workbook.createCellStyle();
        borderedStyle.setBorderTop(BorderStyle.THIN);
        borderedStyle.setBorderBottom(BorderStyle.THIN);
        borderedStyle.setBorderLeft(BorderStyle.THIN);
        borderedStyle.setBorderRight(BorderStyle.THIN);
        borderedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        borderedStyle.setWrapText(true);
        borderedStyle.setFont(normalFont);

        if (danhSach.isEmpty()) {
            Row row = sheet.createRow(startRow);
            Cell cell = row.createCell(0);
            cell.setCellValue("Không có giao dịch nào");

            CellStyle style = workbook.createCellStyle();
            Font redFont = workbook.createFont();
            redFont.setColor(IndexedColors.RED.getIndex());
            redFont.setFontName("Times New Roman");
            redFont.setFontHeightInPoints((short) 14); // 👉 Font đỏ cũng 14pt
            style.setFont(redFont);
            cell.setCellStyle(style);
        } else {
            int stt = 1;
            for (TaiSanTheChap item : danhSach) {
                Row row = sheet.createRow(startRow);

                row.createCell(0).setCellValue(stt++);
                row.createCell(1).setCellValue(item.hoTen);
                row.createCell(4).setCellValue(item.diaChi);
                row.createCell(8).setCellValue(item.soLuong);

                // 👉 Trích xuất và ghi số seri đã lọc
                List<String> seriList = extractSeri(item.seri);
                row.createCell(9).setCellValue(String.join(", ", seriList));

                int[] singleColumns = {0, 8, 9};
                for (int col : singleColumns) {
                    Cell cell = row.getCell(col);
                    if (cell == null) cell = row.createCell(col);
                    cell.setCellStyle(borderedStyle);
                }

                for (int col = 1; col <= 3; col++) {
                    Cell cell = row.getCell(col);
                    if (cell == null) cell = row.createCell(col);
                    cell.setCellStyle(borderedStyle);
                }

                for (int col = 4; col <= 7; col++) {
                    Cell cell = row.getCell(col);
                    if (cell == null) cell = row.createCell(col);
                    cell.setCellStyle(borderedStyle);
                }

                removeOverlappingMergedRegions(sheet, startRow, startRow, 1, 3);
                removeOverlappingMergedRegions(sheet, startRow, startRow, 4, 7);

                sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 1, 3));
                sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 4, 7));

                startRow++;
            }
        }

        // 👉 Đặt độ rộng cột seri để hiển thị đầy đủ
        sheet.setColumnWidth(9, 4000);
    }



    private static void removeOverlappingMergedRegions(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        List<Integer> toRemove = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            boolean rowsOverlap = !(region.getLastRow() < firstRow || region.getFirstRow() > lastRow);
            boolean colsOverlap = !(region.getLastColumn() < firstCol || region.getFirstColumn() > lastCol);
            if (rowsOverlap && colsOverlap) {
                toRemove.add(i);
            }
        }
        for (int i = toRemove.size() - 1; i >= 0; i--) {
            sheet.removeMergedRegion(toRemove.get(i));
        }
    }

    public static String boDau(String text) {
        text = Normalizer.normalize(text, Normalizer.Form.NFD);
        return text.replaceAll("\\p{InCombiningDiacriticalMarks}+", "");
    }
    public static List<String> extractSeri(String input) {
        List<String> result = new ArrayList<>();
        // Biểu thức chính quy cho phép 1–3 chữ cái + tùy chọn khoảng trắng + 6–8 chữ số
        Pattern pattern = Pattern.compile("\\b[\\p{L}]{1,3}\\s?\\d{6,8}\\b");
        Matcher matcher = pattern.matcher(input);
        while (matcher.find()) {
            result.add(matcher.group().trim());
        }
        return result;
    }
}
