package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.sql.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    static class TaiSanTheChap {
        String hoTen;
        String diaChi;
        int soLuong;
        String seri;
        String loaiGiaoDich;
        String tk_no;
        String tk_co;

        @Override
        public String toString() {
            return "TaiSanTheChap{" +
                    "hoTen='" + hoTen + '\'' +
                    ", diaChi='" + diaChi + '\'' +
                    ", soLuong=" + soLuong +
                    ", seri='" + seri + '\'' +
                    ", loaiGiaoDich='" + loaiGiaoDich + '\'' +
                    ", tk_no='" + tk_no + '\'' +
                    ", tk_co='" + tk_co + '\'' +
                    '}';
        }
    }

    public static void main(String[] args) {
        // Tạo scheduler với 1 thread
        ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

        Runnable job = () -> {
            System.out.println("🔄 Job chạy lúc: " + java.time.LocalDateTime.now());

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
                    processDate(fixedDate, conn, danhSachFileDaXuat); // gọi hàm xử lý
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
        };

        // Chạy ngay lần đầu, sau đó lặp lại mỗi 10 phút
        scheduler.scheduleAtFixedRate(job, 0, 60, TimeUnit.SECONDS);
    }


    private static void processDate(String fixedDate, Connection conn, List<String> danhSachFileDaXuat) {
        String sql = """
                WITH Giaodich_Filtered AS (
                    SELECT DISTINCT object_id, ten_loai_giao_dich, ma_tk_no, ma_tk_co
                    FROM vwGiao_Dich
                    WHERE\s
                         CAST(Ngay AS DATE) = ?
                        AND ten_loai_giao_dich IN (N'Xuất tài sản thế chấp', N'Nhập tài sản thế chấp', N'Xuất TS giữ hộ', N'Nhập TS giữ hộ')
                        AND object_id IS NOT NULL
                )
                SELECT\s
                    TSTC.ChuTS_Hoten AS [Họ và tên],
                    TSTC.ChuTS_Diachi AS [Địa chỉ],
                    TSTC.tstc_soluong AS [Số lượng],
                    TSTC.tstc_ten AS [Seri],
                    GD.ten_loai_giao_dich AS [Loại giao dịch],
                    GD.ma_tk_no AS [tk_no],
                    GD.ma_tk_co AS [tk_co]
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
                    item.tk_no = rs.getString("tk_no");
                    item.tk_co = rs.getString("tk_co");
                    danhSach.add(item);
                }
            }
        } catch (SQLException e) {
            System.out.println("❌ Lỗi truy vấn SQL cho ngày " + fixedDate + ":");
            e.printStackTrace();
            return;
        }

        System.out.println("📊 Số bản ghi ngày " + fixedDate + ": " + danhSach.size());
        System.err.println("danh Sach --> " + danhSach);
        // Sau khi đã add hết vào danhSach
        for (TaiSanTheChap item : danhSach) {
            boolean tkNoEmpty = (item.tk_no.trim().isEmpty());
            boolean tkCoEmpty = (item.tk_co.trim().isEmpty());
            if (tkNoEmpty && tkCoEmpty) {
                System.out.println("❌ Bản ghi không hợp lệ:");
                System.out.println(" Họ tên: " + item.hoTen);
                System.out.println(" Địa chỉ: " + item.diaChi);
                System.out.println(" Số lượng: " + item.soLuong);
                System.out.println(" Seri: " + item.seri);
                System.out.println(" Loại giao dịch: " + item.loaiGiaoDich);
                System.out.println(" tk_no: " + item.tk_no);
                System.out.println(" tk_co: " + item.tk_co);
                System.out.println("-----------------------------");
            }
        }
        danhSach.removeIf(item ->
                (item.tk_no.trim().isEmpty()) &&
                        (item.tk_co.trim().isEmpty())
        );
        System.out.println("📊 Số bản ghi ngày " + fixedDate + ": SAU KHI XÓA " + danhSach.size());
        // Phân loại giao dịch
        List<TaiSanTheChap> danhSachXuat = new ArrayList<>();
        List<TaiSanTheChap> danhSachNhap = new ArrayList<>();

        for (TaiSanTheChap item : danhSach) {
            if (item.loaiGiaoDich != null && item.loaiGiaoDich.contains("Xuất")) {
                danhSachXuat.add(item);
            } else {
                danhSachNhap.add(item);
            }
        }

        // Tạo workbook và sheet mới
        Workbook workbook = new XSSFWorkbook();
        Sheet sheetXuat = workbook.createSheet("Xuất TSTC");
        Sheet sheetNhap = workbook.createSheet("Nhập TSTC");

        // Chuyển fixedDate thành LocalDate
        LocalDate ngay = LocalDate.parse(fixedDate, DateTimeFormatter.ofPattern("yyyyMMdd"));
        int nextRowXuat = ghiVanBanCoDinhTren(workbook, sheetXuat, "Xuất", ngay);
        int betweenRowXuat = ghiDuLieu(workbook, sheetXuat, danhSachXuat, nextRowXuat);
        int endRowXuat = ghiVanBanCoDinhDuoi(workbook, sheetXuat, betweenRowXuat, "Xuất", ngay);

        int nextRowNhap = ghiVanBanCoDinhTren(workbook, sheetNhap, "Nhập", ngay);
        int betweenNhap = ghiDuLieu(workbook, sheetNhap, danhSachNhap, nextRowNhap);
        int endRowNhap = ghiVanBanCoDinhDuoi(workbook, sheetNhap, betweenNhap, "Nhập", ngay);
// 👉 Đặt vùng in sau khi đã ghi xong tất cả
        setupPrintA4(workbook, sheetXuat, 0, 4, 0, endRowXuat - 1);
        setupPrintA4(workbook, sheetNhap, 0, 4, 0, endRowNhap - 1);

        // Tạo thư mục đầu ra
        try {
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
            workbook.close();
        } catch (Exception e) {
            System.out.println("❌ Lỗi xử lý file Excel cho ngày " + fixedDate + ":");
            e.printStackTrace();
        }
    }


    private static int ghiVanBanCoDinhTren(Workbook workbook, Sheet sheet, String isXuatNhap, LocalDate ngay) {
        int currentRow = 0;

        // Font đậm
        Font font = workbook.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeightInPoints((short) 13);
        font.setBold(true);

        //Font thường
        Font fontNormal = workbook.createFont();
        fontNormal.setFontName("Times New Roman");
        fontNormal.setFontHeightInPoints((short) 13);
        fontNormal.setBold(false);

        CellStyle normalLeftStyle = workbook.createCellStyle();
        normalLeftStyle.setFont(fontNormal);
        normalLeftStyle.setAlignment(HorizontalAlignment.LEFT);
        normalLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle normalCenterStyle = workbook.createCellStyle();
        normalCenterStyle.setFont(fontNormal);
        normalCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        normalCenterStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        CellStyle boldCenterStyle = workbook.createCellStyle();
        boldCenterStyle.setFont(font);
        boldCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        boldCenterStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldCenterStyle.setWrapText(true);

        // Font nghiêng
        Font italicFont = workbook.createFont();
        italicFont.setFontName("Times New Roman");
        italicFont.setFontHeightInPoints((short) 13);
        italicFont.setItalic(true);

        CellStyle leftStyle = workbook.createCellStyle();
        leftStyle.setFont(font);
        leftStyle.setAlignment(HorizontalAlignment.LEFT);
        leftStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle italicLeftStyle = workbook.createCellStyle();
        italicLeftStyle.setFont(italicFont);
        italicLeftStyle.setAlignment(HorizontalAlignment.LEFT);
        italicLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        italicLeftStyle.setWrapText(true);

        CellStyle italicRightStyle = workbook.createCellStyle();
        italicRightStyle.setFont(italicFont);
        italicRightStyle.setAlignment(HorizontalAlignment.CENTER);
        italicRightStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        italicRightStyle.setWrapText(true);

        currentRow = writeHeader(workbook, sheet, currentRow, boldCenterStyle);

        currentRow++;
        // Các dòng căn giữa + đậm
        String[] centeredLines = {
                "QUYẾT ĐỊNH",
                "\"V/v " + isXuatNhap.toLowerCase() + " kho tài sản thế chấp, cầm cố\"",
        };
        currentRow = writeLeftNormalLines(sheet, currentRow, centeredLines, boldCenterStyle, 1, 5);
        // Hai dòng nghiêng
        String[] italicLines = {
                "- Căn cứ vào quy chế kho quỹ của Quỹ tín dụng nhân dân Thái Học",
                "- Căn cứ vào tình hình hoạt động của Quỹ tín dụng Thái Học"
        };
        for (String line : italicLines) {
            Row r = sheet.createRow(currentRow++);
            Cell c = r.createCell(0);
            c.setCellValue(line);
            c.setCellStyle(italicLeftStyle);
            mergeSafe(sheet, new CellRangeAddress(r.getRowNum(), r.getRowNum(), 0, 4));
            r.setHeightInPoints(22);
        }
        currentRow++;
        String[] centeredLines2 = {
                "BAN ĐIỀU HÀNH QTD THÁI HỌC",
                "QUYẾT ĐỊNH " + isXuatNhap.toUpperCase() + " KHO",
        };
        currentRow = writeLeftNormalLines(sheet, currentRow, centeredLines2, boldCenterStyle, 1, 5);
        currentRow++;
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "I. " + isXuatNhap + " kho tài sản thế chấp, cầm cố của khách hàng:",
                leftStyle, 0, 6);

        currentRow = writeLeftNormalLines(sheet, currentRow,
                new String[]{"- " + isXuatNhap + " kho tài sản thế chấp, cầm cố của khách hàng (có bảng kê kèm theo)"},
                normalLeftStyle, 0, 6);


        currentRow = writeLeftBoltLine(sheet, currentRow,
                "II. Người chịu trách nhiệm vận chuyển số tài sản trên:",
                leftStyle, 0, 6);
        String[] row151617 = {
                "1. Bà: Phùng Thị Loan - Giám đốc",
                "2. Ông: Vũ Đình Kiên - Kế toán",
                "3. Ông: Nguyễn Đồng Chính - Thủ quỹ (thủ kho)"
        };

        currentRow = writeLeftNormalLines(sheet, currentRow, row151617, normalLeftStyle, 0, 6);
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "III. Ông (bà) kế toán trưởng, thủ quỹ và các ông (bà) có tên trên:",
                leftStyle, 0, 6);
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "chịu trách nhiệm quyết định thi hành này",
                leftStyle, 0, 6);
        // Ngày tháng năm
        String ngayThangNam = String.format("Chu Văn An, ngày %02d tháng %02d năm %d",
                ngay.getDayOfMonth(), ngay.getMonthValue(), ngay.getYear());
        Row rDate = sheet.createRow(currentRow++);
        Cell cDate = rDate.createCell(2);
        cDate.setCellValue(ngayThangNam);
        cDate.setCellStyle(italicRightStyle);
        mergeSafe(sheet, new CellRangeAddress(rDate.getRowNum(), rDate.getRowNum(), 2, 5));
        rDate.setHeightInPoints(22);
        String[] chuKyGD = {
                "T/M QTD THÁI HỌC",
                "GIÁM ĐỐC",
        };
        currentRow = writeLeftNormalLines(sheet, currentRow, chuKyGD, boldCenterStyle, 2, 5);
        currentRow += 5;
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "Phùng Thị Loan",
                normalCenterStyle, 2, 5);
        // Thiết lập khổ in A4
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE);
        printSetup.setLandscape(false);
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setAutobreaks(true);

        sheet.setMargin(Sheet.LeftMargin, 0.3);
        sheet.setMargin(Sheet.RightMargin, 0.3);
        sheet.setMargin(Sheet.TopMargin, 0.5);
        sheet.setMargin(Sheet.BottomMargin, 0.5);

        // Đặt độ rộng cột
        for (int i = 0; i <= 4; i++) {
            sheet.setColumnWidth(i, 6000);
        }
        sheet.setRowBreak(currentRow);
        currentRow += 2;
        currentRow = writeHeader(workbook, sheet, currentRow, boldCenterStyle);
        currentRow++;
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "BẢNG KÊ " + isXuatNhap.toUpperCase() + " KHO",
                boldCenterStyle, 1, 4);
        // Ngày tháng năm
        String ngayThangNam2 = String.format("Ngày %02d tháng %02d năm %d",
                ngay.getDayOfMonth(), ngay.getMonthValue(), ngay.getYear());
        Row rDate2 = sheet.createRow(currentRow++);
        Cell cDate2 = rDate2.createCell(1);
        cDate2.setCellValue(ngayThangNam2);
        cDate2.setCellStyle(normalCenterStyle);
        mergeSafe(sheet, new CellRangeAddress(rDate2.getRowNum(), rDate2.getRowNum(), 1, 4));
        rDate2.setHeightInPoints(22);

        currentRow = writeLeftNormalLines(sheet, currentRow,
                new String[]{"- " + isXuatNhap + " kho tài sản thế chấp, cầm cố của khách hàng"},
                normalLeftStyle, 0, 6);
        setupPrintA4(workbook, sheet, 0, 4, 0, currentRow - 1);

        return currentRow;
    }

    private static int ghiVanBanCoDinhDuoi(Workbook workbook, Sheet sheet, int startRow, String isXuatNhap, LocalDate ngay) {
        startRow++;
        int currentRow = startRow;

        // Font thường
        Font fontNormal = workbook.createFont();
        fontNormal.setFontName("Times New Roman");
        fontNormal.setFontHeightInPoints((short) 13);

        // Font chữ ký
        Font chuKy = workbook.createFont();
        chuKy.setFontName("Times New Roman");
        chuKy.setFontHeightInPoints((short) 12);

        // Font đậm
        Font fontBold = workbook.createFont();
        fontBold.setFontName("Times New Roman");
        fontBold.setFontHeightInPoints((short) 13);
        fontBold.setBold(true);
        // Font nghiêng
        Font italicFont = workbook.createFont();
        italicFont.setFontName("Times New Roman");
        italicFont.setFontHeightInPoints((short) 13);
        italicFont.setItalic(true);

        CellStyle boldLeftStyle = workbook.createCellStyle();
        boldLeftStyle.setFont(fontBold);
        boldLeftStyle.setAlignment(HorizontalAlignment.LEFT);
        boldLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle normalLeftStyle = workbook.createCellStyle();
        normalLeftStyle.setFont(fontNormal);
        normalLeftStyle.setAlignment(HorizontalAlignment.LEFT);
        normalLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle italicLeftStyle = workbook.createCellStyle();
        italicLeftStyle.setFont(italicFont);
        italicLeftStyle.setAlignment(HorizontalAlignment.LEFT);
        italicLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        italicLeftStyle.setWrapText(true);

        CellStyle italicCenterStyle = workbook.createCellStyle();
        italicCenterStyle.setFont(italicFont);
        italicCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        italicCenterStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        italicCenterStyle.setWrapText(true);
        // Style căn giữa
        CellStyle centerStyle = workbook.createCellStyle();
        centerStyle.setFont(fontNormal);
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Style căn giữa
        CellStyle chuKyStyle = workbook.createCellStyle();
        chuKyStyle.setFont(chuKy);
        chuKyStyle.setAlignment(HorizontalAlignment.CENTER);
        chuKyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle boldCenterStyle = workbook.createCellStyle();
        boldCenterStyle.setFont(fontBold);
        boldCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        boldCenterStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        // Dòng tiêu đề các chức danh
        Row rowChucDanh = sheet.createRow(currentRow++);
        String[] chucDanh = {"THỦ KHO", "KẾ TOÁN", "GIÁM ĐỐC"};
        for (int i = 0; i < chucDanh.length; i++) {
            Cell cell = rowChucDanh.createCell(i * 2);
            cell.setCellValue(chucDanh[i]);
            cell.setCellStyle(boldCenterStyle);
            sheet.addMergedRegion(new CellRangeAddress(rowChucDanh.getRowNum(), rowChucDanh.getRowNum(), i * 2, i * 2 + 1));
        }

        // Dòng ghi chú ký tên
        Row rowGhiChu = sheet.createRow(currentRow++);
        String[] ghiChu = {"(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)"};
        for (int i = 0; i < ghiChu.length; i++) {
            Cell cell = rowGhiChu.createCell(i * 2);
            cell.setCellValue(ghiChu[i]);
            cell.setCellStyle(chuKyStyle);
            sheet.addMergedRegion(new CellRangeAddress(rowGhiChu.getRowNum(), rowGhiChu.getRowNum(), i * 2, i * 2 + 1));
        }

        // Dòng tên người ký
        currentRow += 4; // tạo khoảng trống cho chữ ký
        Row rowTen = sheet.createRow(currentRow++);
        String[] tenNguoiKy = {"Nguyễn Đồng Chính", "Vũ Đình Kiên", "Phùng Thị Loan"};
        for (int i = 0; i < tenNguoiKy.length; i++) {
            Cell cell = rowTen.createCell(i * 2);
            cell.setCellValue(tenNguoiKy[i]);
            cell.setCellStyle(centerStyle);
            sheet.addMergedRegion(new CellRangeAddress(rowTen.getRowNum(), rowTen.getRowNum(), i * 2, i * 2 + 1));
        }
        sheet.setRowBreak(currentRow);
        currentRow += 2;
        currentRow = writeHeaderDuoi(workbook, sheet, currentRow, boldCenterStyle);
        currentRow++;
        String[] centeredLines = {
                "QUYẾT ĐỊNH",
                "\"V/v " + isXuatNhap.toLowerCase() + " kho hòm tôn bảo quản tiền mặt, giấy tờ có giá\"",
        };

        currentRow = writeLeftNormalLines(sheet, currentRow, centeredLines, boldCenterStyle, 1, 5);
        // Hai dòng nghiêng
        String[] italicLines = {
                "- Căn cứ vào quy chế kho quỹ của Quỹ tín dụng nhân dân Thái Học",
                "- Căn cứ vào tình hình hoạt động của Quỹ tín dụng Thái Học"
        };
        for (String line : italicLines) {
            Row r = sheet.createRow(currentRow++);
            Cell c = r.createCell(0);
            c.setCellValue(line);
            c.setCellStyle(italicLeftStyle);
            mergeSafe(sheet, new CellRangeAddress(r.getRowNum(), r.getRowNum(), 0, 4));
            r.setHeightInPoints(22);
        }
        currentRow++;
        String[] centeredLines2 = {
                "BAN ĐIỀU HÀNH QTD THÁI HỌC",
                "QUYẾT ĐỊNH " + isXuatNhap.toUpperCase() + " KHO",
        };
        currentRow = writeLeftNormalLines(sheet, currentRow, centeredLines2, boldCenterStyle, 1, 5);
        currentRow++;
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "I. " + isXuatNhap + " kho tiền mặt, các loại giấy tờ có giá cụ thể như sau:",
                boldLeftStyle, 0, 6);
        currentRow = writeLeftNormalLines(sheet, currentRow,
                new String[]{"- " + isXuatNhap + " kho 01 hòm tôn bảo quản tiền mặt, giấy tờ có giá trong giờ nghỉ trưa"},
                normalLeftStyle, 0, 6);
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "II. Người chịu trách nhiệm vận chuyển số tài sản trên:",
                boldLeftStyle, 0, 6);
        String[] row151617 = {
                "1. Bà: Phùng Thị Loan - Giám đốc",
                "2. Ông: Vũ Đình Kiên - Kế toán",
                "3. Ông: Nguyễn Đồng Chính - Thủ quỹ (thủ kho)"
        };
        currentRow = writeLeftNormalLines(sheet, currentRow, row151617, normalLeftStyle, 0, 6);
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "III. Ông (bà) kế toán trưởng, thủ quỹ và các ông (bà) có tên trên:",
                boldLeftStyle, 0, 6);
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "chịu trách nhiệm quyết định thi hành này",
                boldLeftStyle, 0, 6);
        // Ngày tháng năm
        String ngayThangNam = String.format("Chu Văn An, ngày %02d tháng %02d năm %d",
                ngay.getDayOfMonth(), ngay.getMonthValue(), ngay.getYear());
        Row rDate = sheet.createRow(currentRow++);
        Cell cDate = rDate.createCell(2);
        cDate.setCellValue(ngayThangNam);
        cDate.setCellStyle(italicCenterStyle);
        mergeSafe(sheet, new CellRangeAddress(rDate.getRowNum(), rDate.getRowNum(), 2, 5));
        rDate.setHeightInPoints(22);
        String[] chuKyGD = {
                "T/M QTD THÁI HỌC",
                "GIÁM ĐỐC",
        };
        currentRow = writeLeftNormalLines(sheet, currentRow, chuKyGD, boldCenterStyle, 2, 5);
        currentRow += 5;
        currentRow = writeLeftBoltLine(sheet, currentRow,
                "Phùng Thị Loan",
                centerStyle, 2, 5);
        setupPrintA4(workbook, sheet, 0, 6, 0, currentRow - 1);
        return currentRow;
    }

    private static int writeHeaderDuoi(Workbook workbook, Sheet sheet, int currentRow, CellStyle boldCenterStyle) {
        Row row0 = sheet.createRow(currentRow++);
        row0.setHeightInPoints(22);

        CellStyle leftBoldStyle = workbook.createCellStyle();
        leftBoldStyle.cloneStyleFrom(boldCenterStyle);
        leftBoldStyle.setAlignment(HorizontalAlignment.LEFT);

        Cell cellLeft = row0.createCell(0);
        cellLeft.setCellValue("QTDND THÁI HỌC");
        cellLeft.setCellStyle(leftBoldStyle);
        mergeSafe(sheet, new CellRangeAddress(row0.getRowNum(), row0.getRowNum(), 0, 1));

        Cell cellRight = row0.createCell(2);
        cellRight.setCellValue("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
        cellRight.setCellStyle(boldCenterStyle);
        mergeSafe(sheet, new CellRangeAddress(row0.getRowNum(), row0.getRowNum(), 2, 6));

        // Không gọi setColumnWidth ở đây

        Font underlineFont = workbook.createFont();
        underlineFont.setFontName("Times New Roman");
        underlineFont.setFontHeightInPoints((short) 13);
        underlineFont.setBold(true);
        underlineFont.setUnderline(Font.U_SINGLE);

        CellStyle sloganStyle = workbook.createCellStyle();
        sloganStyle.setFont(underlineFont);
        sloganStyle.setAlignment(HorizontalAlignment.CENTER);
        sloganStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Row row1 = sheet.createRow(currentRow++);
        row1.setHeightInPoints(22);
        Cell cellSlogan = row1.createCell(2);
        cellSlogan.setCellValue("Độc lập – Tự do – Hạnh phúc");
        cellSlogan.setCellStyle(sloganStyle);
        mergeSafe(sheet, new CellRangeAddress(row1.getRowNum(), row1.getRowNum(), 2, 6));

        return currentRow;
    }


    /**
     * Ghi phần tiêu đề: bên trái + quốc hiệu + khẩu hiệu gạch chân
     *
     * @param workbook        Workbook hiện tại
     * @param sheet           Sheet cần ghi
     * @param currentRow      dòng bắt đầu
     * @param boldCenterStyle Style căn giữa + đậm
     * @return chỉ số dòng tiếp theo
     */
    private static int writeHeader(Workbook workbook, Sheet sheet, int currentRow,
                                   CellStyle boldCenterStyle) {
        // Row 0: bên trái + quốc hiệu
        Row row0 = sheet.createRow(currentRow);
        row0.setHeightInPoints(22);

        // Style trái (đậm + căn trái)
        CellStyle leftBoldStyle = workbook.createCellStyle();
        leftBoldStyle.cloneStyleFrom(boldCenterStyle);
        leftBoldStyle.setAlignment(HorizontalAlignment.LEFT);
        String text = "QTDND THÁI HỌC";
        Cell cellLeft = row0.createCell(0);
        cellLeft.setCellValue(text);
        cellLeft.setCellStyle(leftBoldStyle);
        mergeSafe(sheet, new CellRangeAddress(row0.getRowNum(), row0.getRowNum(), 0, 1));
        Cell cellRight = row0.createCell(2);
        cellRight.setCellValue("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
        cellRight.setCellStyle(boldCenterStyle);
        mergeSafe(sheet, new CellRangeAddress(row0.getRowNum(), row0.getRowNum(), 2, 6));
        for (int i = 2; i <= 6; i++) {
            sheet.setColumnWidth(i, 1000); // hoặc 4800 nếu cần cân đối
        }
        // Tăng dòng sau khi tạo row0
        currentRow++;

        // Tạo font gạch chân cho khẩu hiệu
        Font underlineFont = workbook.createFont();
        underlineFont.setFontName("Times New Roman");
        underlineFont.setFontHeightInPoints((short) 13);
        underlineFont.setBold(true);
        underlineFont.setUnderline(Font.U_SINGLE);

        CellStyle sloganStyle = workbook.createCellStyle();
        sloganStyle.setFont(underlineFont);
        sloganStyle.setAlignment(HorizontalAlignment.CENTER);
        sloganStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Row 1: khẩu hiệu (ở dòng tiếp theo)
        Row row1 = sheet.createRow(currentRow);
        row1.setHeightInPoints(22);
        Cell cellSlogan = row1.createCell(2);
        cellSlogan.setCellValue("Độc lập – Tự do – Hạnh phúc");
        cellSlogan.setCellStyle(sloganStyle);
        mergeSafe(sheet, new CellRangeAddress(row1.getRowNum(), row1.getRowNum(), 2, 6));

        // Tăng dòng sau khi tạo row1
        currentRow++;

        return currentRow;
    }


    /**
     * Ghi một hoặc nhiều dòng văn bản vào sheet với style và merge vùng
     *
     * @param sheet      Sheet cần ghi
     * @param currentRow chỉ số dòng hiện tại
     * @param lines      mảng các chuỗi cần ghi (có thể 1 hoặc nhiều phần tử)
     * @param style      CellStyle áp dụng
     * @param firstCol   cột bắt đầu merge
     * @param lastCol    cột kết thúc merge
     * @return chỉ số dòng tiếp theo
     */
    private static int writeLeftNormalLines(Sheet sheet, int currentRow, String[] lines,
                                            CellStyle style, int firstCol, int lastCol) {
        for (String line : lines) {
            Row row = sheet.createRow(currentRow++);
            Cell cell = row.createCell(firstCol);
            cell.setCellValue(line);
            cell.setCellStyle(style);

            mergeSafe(sheet, new CellRangeAddress(row.getRowNum(), row.getRowNum(), firstCol, lastCol));
            row.setHeightInPoints(22);
        }
        return currentRow;
    }


    /**
     * Hàm merge an toàn: chỉ merge nếu chưa tồn tại vùng đó
     */
    private static void mergeSafe(Sheet sheet, CellRangeAddress region) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            if (sheet.getMergedRegion(i).equals(region)) {
                return; // đã tồn tại, bỏ qua
            }
        }
        sheet.addMergedRegion(region);
    }

    private static int writeLeftBoltLine(Sheet sheet, int currentRow, String text,
                                         CellStyle style, int firstCol, int lastCol) {
        Row row = sheet.createRow(currentRow++);
        Cell cell = row.createCell(firstCol);
        cell.setCellValue(text);
        cell.setCellStyle(style);

        mergeSafe(sheet, new CellRangeAddress(row.getRowNum(), row.getRowNum(), firstCol, lastCol));
        row.setHeightInPoints(22);

        return currentRow;
    }

    private static int ghiDuLieu(Workbook workbook, Sheet sheet, List<TaiSanTheChap> danhSach, int startRow) {
        int currentRow = startRow;

        // Font thường
        Font normalFont = workbook.createFont();
        normalFont.setFontName("Times New Roman");
        normalFont.setFontHeightInPoints((short) 14);

        // Font in đậm cho header
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setFontHeightInPoints((short) 14);
        boldFont.setBold(true);

        // Style cho dữ liệu có border + wrap text
        CellStyle borderedStyle = workbook.createCellStyle();
        borderedStyle.setBorderTop(BorderStyle.THIN);
        borderedStyle.setBorderBottom(BorderStyle.THIN);
        borderedStyle.setBorderLeft(BorderStyle.THIN);
        borderedStyle.setBorderRight(BorderStyle.THIN);
        borderedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        borderedStyle.setFont(normalFont);
        borderedStyle.setWrapText(true); // Cho phép xuống dòng

        // Style cho header: border + bold + căn giữa
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFont(boldFont);

        // 👉 Header bảng
        Row header = sheet.createRow(currentRow++);
        header.setHeightInPoints(20);
        header.createCell(0).setCellValue("STT");
        header.createCell(1).setCellValue("Họ và tên");
        header.createCell(2).setCellValue("Địa chỉ");
        header.createCell(3).setCellValue("Số lượng");
        header.createCell(4).setCellValue("Seri");

        for (int col = 0; col <= 4; col++) {
            header.getCell(col).setCellStyle(headerStyle);
        }

        // 👉 Dữ liệu động
        if (danhSach.isEmpty()) {
            Row row = sheet.createRow(currentRow++);
            Cell cell = row.createCell(0);
            cell.setCellValue("Không có giao dịch nào");

            CellStyle redStyle = workbook.createCellStyle();
            Font redFont = workbook.createFont();
            redFont.setColor(IndexedColors.RED.getIndex());
            redFont.setFontName("Times New Roman");
            redFont.setFontHeightInPoints((short) 14);
            redStyle.setFont(redFont);
            cell.setCellStyle(redStyle);
        } else {
            int stt = 1;
            for (TaiSanTheChap item : danhSach) {
                String diaChiLoc = extractFirstAddressPart(item.diaChi);
                List<String> seriLoc = extractSeri(item.seri);
                String seriChuoi = seriLoc.isEmpty() ? item.seri : String.join("\n", seriLoc);
                int lineCount = seriChuoi.split("\n").length;

                Row row = sheet.createRow(currentRow++);
                row.setHeightInPoints(lineCount * 15); // Điều chỉnh chiều cao dòng theo số dòng

                row.createCell(0).setCellValue(stt++);
                row.createCell(1).setCellValue(item.hoTen);
                row.createCell(2).setCellValue(diaChiLoc);
                row.createCell(3).setCellValue(item.soLuong);
                row.createCell(4).setCellValue(seriChuoi);

                for (int col = 0; col <= 4; col++) {
                    row.getCell(col).setCellStyle(borderedStyle);
                }
            }
        }

        // 👉 Đặt độ rộng cột cố định cho các cột khác
        sheet.setColumnWidth(0, 1500);
        sheet.setColumnWidth(2, 6000);
        sheet.setColumnWidth(3, 3200);
        sheet.setColumnWidth(4, 5000);

        // 👉 Tự động fit cột "Họ và tên"
        sheet.autoSizeColumn(1);

        // 👉 Giới hạn để không vượt khổ A4
        int maxWidth = 8000; // ~70-80 ký tự Times New Roman 14pt
        if (sheet.getColumnWidth(1) > maxWidth) {
            sheet.setColumnWidth(1, maxWidth);
        }

        return currentRow;
    }


    /**
     * Thiết lập khổ in A4 và căn giữa cho toàn bộ sheet
     *
     * @param workbook Workbook chứa sheet
     * @param sheet    Sheet cần thiết lập
     * @param firstCol cột bắt đầu vùng in
     * @param lastCol  cột kết thúc vùng in
     * @param firstRow dòng bắt đầu vùng in
     * @param lastRow  dòng kết thúc vùng in
     */
    private static void setupPrintA4(Workbook workbook, Sheet sheet,
                                     int firstCol, int lastCol,
                                     int firstRow, int lastRow) {
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE);
        printSetup.setLandscape(false); // true nếu muốn in ngang
        // 👉 Fit to page
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setAutobreaks(true);

        sheet.setHorizontallyCenter(true); // căn giữa ngang
        // sheet.setVerticallyCenter(true); // nếu muốn căn giữa dọc

        // Đặt vùng in
        workbook.setPrintArea(
                workbook.getSheetIndex(sheet),
                firstCol, lastCol,
                firstRow, lastRow
        );

        // Margin
        sheet.setMargin(Sheet.LeftMargin, 0.1);
        sheet.setMargin(Sheet.RightMargin, 0.1);
        sheet.setMargin(Sheet.TopMargin, 0.8);
        sheet.setMargin(Sheet.BottomMargin, 0.5);
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

    private static String extractFirstAddressPart(String diaChi) {
        if (diaChi == null || diaChi.isBlank()) return diaChi;
        // Tách theo dấu '-' hoặc ','
        String[] parts = diaChi.split("[-,]");
        return parts[0].trim();
    }


}
