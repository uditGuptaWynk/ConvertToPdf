package org.example;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Worksheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import com.aspose.cells.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) throws Exception {

        String home = System.getProperty("user.home");

        FileInputStream fisCode = new FileInputStream(new File(home + "/Desktop/ArtistTest/ArtistPaymentStatus.xlsx"));
        XSSFWorkbook wbReadCode = new XSSFWorkbook(fisCode);
        Sheet sheet2Code = wbReadCode.getSheetAt(0);
        int rowCountCode = sheet2Code.getLastRowNum();
        System.out.println(rowCountCode);

//        Read artist data
        FileInputStream fis = new FileInputStream(new File(home + "/Desktop/ArtistTest/ArtistData.xlsx"));
        XSSFWorkbook wbRead = new XSSFWorkbook(fis);
        Sheet sheet2 = wbRead.getSheetAt(0);
        int rowCount = sheet2.getLastRowNum();
        System.out.println(rowCount+" Count");
//        rowCount = 100;
        for(int p=1; p<rowCountCode+1; p++){
            System.out.println(p+"Count");
            String artistIdCode = sheet2Code.getRow(p).getCell(CellReference.convertColStringToIndex("E")).getStringCellValue();
            if(artistIdCode.equals("")){
                break;
            }
            artistIdCode = sheet2Code.getRow(p).getCell(CellReference.convertColStringToIndex("E")).getStringCellValue();
            Double vendorCode = sheet2Code.getRow(p).getCell(CellReference.convertColStringToIndex("F")).getNumericCellValue();

            for (int j = 1; j < rowCount; j++) {
            Boolean pdfCreate = true;
//        int j=1;
            String artistName = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("E")).getStringCellValue();
            String artistId = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("D")).getStringCellValue();

            if(artistId.equals(artistIdCode)) {
//        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
                long currentTimeMillis = System.currentTimeMillis();
                String invoiceNo = "WYNK/" + currentTimeMillis;
                String month = String.valueOf(sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("A")).getLocalDateTimeCellValue().toLocalDate());
                double descNum = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("B")).getNumericCellValue();
                double rate = 0.03;
                double amount = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("C")).getNumericCellValue();
                ArrayList descNumList = new ArrayList<Double>();
                ArrayList monthList = new ArrayList<String>();
                ArrayList amountList = new ArrayList<Double>();
                System.out.println(descNumList.size());
                DataFormatter dataFormatter = new DataFormatter();

                String address = "";
                String accountName = "";
                String accountNo = null;
                String bankName = "";
                String IFSCCode = "";
                String PANNo = "";
                try {
                    address = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("J")).getStringCellValue();
                    accountName = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("M")).getStringCellValue();
                    accountNo = (long) sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("O")).getNumericCellValue() +"";
                    bankName = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("N")).getStringCellValue();
                    IFSCCode = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("P")).getStringCellValue();
                    PANNo = sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("Q")).getStringCellValue();
                } catch (IllegalStateException e) {
//                pdfCreate =false;
                } catch (Exception e){
                    System.out.println("----------------------------------------");
                    System.out.println(artistId);
                    System.out.println(artistName);
                    System.out.println(j);
                    System.out.println("----------------------------------------");
                    pdfCreate =false;
                }

                if (!pdfCreate || artistId.equals(null) || artistId.equals("#N/A") || artistId.equals("") || month.equals(null) || month.equals("#N/A") || month.equals("") ||
                        accountName.equals(null) || accountName.equals("#N/A") || accountName.equals("") || bankName.equals(null) || bankName.equals("#N/A") || bankName.equals("") || bankName.equals("0") ||
                        IFSCCode.equals(null) || IFSCCode.equals("#N/A") || IFSCCode.equals("") || IFSCCode.equals("0") || PANNo.equals(null) || PANNo.equals("#N/A") || PANNo.equals("") || PANNo.equals("0")) {
                    pdfCreate = false;
                }
                System.out.println(artistId + " Udit");

                if (pdfCreate) {

                    for (int i = j + 1; i < rowCount+1; i++) {
                        String artistId1 = sheet2.getRow(i).getCell(CellReference.convertColStringToIndex("D")).getStringCellValue();
                        if (artistId1.equals(artistId)) {
                            descNumList.add(sheet2.getRow(i).getCell(CellReference.convertColStringToIndex("B")).getNumericCellValue());
                            monthList.add(sheet2.getRow(i).getCell(CellReference.convertColStringToIndex("A")).getLocalDateTimeCellValue().toLocalDate());
                            amountList.add(sheet2.getRow(i).getCell(CellReference.convertColStringToIndex("C")).getNumericCellValue());
                        }
                    }
                    System.out.println(monthList);
                    System.out.println(descNumList);
//        sheet2.getRow(j).getCell(CellReference.convertColStringToIndex("M")).setCellType(Cell.CELL_TYPE_STRING);


                    Row row;
//        Create excel
                    XSSFWorkbook wb = new XSSFWorkbook();
                    Sheet sheet1 = wb.createSheet("new sheet");
                    FileOutputStream fileOut = new FileOutputStream(home + "/Desktop/ArtistTest/statement.xlsx");

                    XSSFCellStyle style = wb.createCellStyle();
                    XSSFCellStyle tableStyle = wb.createCellStyle();
                    XSSFCellStyle tableStyleBold = wb.createCellStyle();
                    XSSFFont font1 = wb.createFont();
                    font1.setBold(true);
                    font1.setFontHeight(13);
                    CellStyle cellStyle = wb.createCellStyle();
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    cellStyle.setFont(font1);
                    CellStyle cellStyleCenter = wb.createCellStyle();
                    cellStyleCenter.setAlignment(HorizontalAlignment.CENTER);
                    CellStyle cellStyleLeft = wb.createCellStyle();
                    cellStyleLeft.setAlignment(HorizontalAlignment.LEFT);
                    XSSFFont font = wb.createFont();
                    font.setBold(true);
                    style.setFont(font);
                    tableStyleBold.setFont(font);
                    tableStyle.setBorderTop(BorderStyle.HAIR);
                    tableStyle.setBorderBottom(BorderStyle.HAIR);
                    tableStyle.setBorderLeft(BorderStyle.HAIR);
                    tableStyle.setBorderRight(BorderStyle.HAIR);
                    tableStyle.setAlignment(HorizontalAlignment.LEFT);
                    tableStyleBold.setBorderTop(BorderStyle.HAIR);
                    tableStyleBold.setBorderBottom(BorderStyle.HAIR);
                    tableStyleBold.setBorderLeft(BorderStyle.HAIR);
                    tableStyleBold.setBorderRight(BorderStyle.HAIR);


                    sheet1.createRow(0).createCell(1).setCellValue("Invoice");
                    sheet1.getRow(0).getCell(1).setCellStyle(cellStyle);
                    sheet1.createRow(2).createCell(1).setCellValue(artistName);
                    sheet1.getRow(2).getCell(1).setCellStyle(style);
                    sheet1.createRow(4).createCell(1).setCellValue(address);
                    sheet1.getRow(4).getCell(1).setCellStyle(style);
                    row = sheet1.createRow(6);
                    row.createCell(1).setCellValue("Invoice Number");
                    row.createCell(2).setCellValue(invoiceNo);
                    sheet1.getRow(6).getCell(2).setCellStyle(cellStyleLeft);
                    row = sheet1.createRow(7);
                    row.createCell(1).setCellValue("Supplier Code");
                    row.createCell(2).setCellValue(vendorCode);
                    row = sheet1.createRow(8);
                    row.createCell(1).setCellValue("Date");
                    long millis = System.currentTimeMillis();
                    java.sql.Date date = new java.sql.Date(millis);
                    row.createCell(2).setCellValue(date.toString());
                    sheet1.getRow(6).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(6).getCell(2).setCellStyle(tableStyle);
                    sheet1.getRow(7).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(7).getCell(2).setCellStyle(tableStyle);
                    sheet1.getRow(8).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(8).getCell(2).setCellStyle(tableStyle);
                    sheet1.createRow(10).createCell(1).setCellValue("To,");
                    sheet1.createRow(11).createCell(1).setCellValue("AIRTEL DIGITAL LIMITED");
                    sheet1.createRow(12).createCell(1).setCellValue("Plot 16, Airtel Centre, Udyog Vihar , Phase IV, Gurgaon Haryana, 122015");
                    sheet1.createRow(13).createCell(1).setCellValue("GST - 06AABCW6047M1ZZ");
                    if (monthList.size() > 0) {
                        sheet1.createRow(15).createCell(1).setCellValue("Period of Service: " + month + " - " + monthList.get(monthList.size() - 1));
                    } else {
                        sheet1.createRow(15).createCell(1).setCellValue("Period of Service: " + month + " - " + month);
                    }
                    sheet1.getRow(15).getCell(1).setCellStyle(style);
                    row = sheet1.createRow(16);
                    row.createCell(1).setCellValue("Details - Month");
                    row.createCell(2).setCellValue("Description Numbers");
                    row.createCell(3).setCellValue("Rate");
                    row.createCell(4).setCellValue("Amount");
                    sheet1.getRow(16).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(16).getCell(2).setCellStyle(tableStyleBold);
                    sheet1.getRow(16).getCell(3).setCellStyle(tableStyleBold);
                    sheet1.getRow(16).getCell(4).setCellStyle(tableStyleBold);
//        XSSFCellStyle style1 = wb.createCellStyle();


                    row = sheet1.createRow(17);
                    row.createCell(1).setCellValue(month);
                    row.createCell(2).setCellValue(descNum);
                    row.createCell(3).setCellValue(rate);
                    row.createCell(4).setCellValue(amount);
                    sheet1.getRow(17).getCell(1).setCellStyle(tableStyle);
                    sheet1.getRow(17).getCell(2).setCellStyle(tableStyle);
                    sheet1.getRow(17).getCell(3).setCellStyle(tableStyle);
                    sheet1.getRow(17).getCell(4).setCellStyle(tableStyle);

                    int count = 18;
                    double num;
                    double total = amount;
                    LocalDateTime localDateTime;
                    LocalDate localDate;
                    System.out.println(monthList.size()+" size");
                    if (monthList.size() > 0) {
                        for (int i = 0; i < monthList.size(); i++) {
                            num = (double) descNumList.get(i);
                            row = sheet1.createRow(count);
                            localDate = (LocalDate) monthList.get(i);
                            row.createCell(1).setCellValue(String.valueOf(localDate));
                            row.createCell(2).setCellValue((Double) descNumList.get(i));
                            row.createCell(3).setCellValue(rate);
                            row.createCell(4).setCellValue((Double) amountList.get(i));
                            total = total + ((Double) amountList.get(i));
                            sheet1.getRow(count).getCell(1).setCellStyle(tableStyle);
                            sheet1.getRow(count).getCell(2).setCellStyle(tableStyle);
                            sheet1.getRow(count).getCell(3).setCellStyle(tableStyle);
                            sheet1.getRow(count).getCell(4).setCellStyle(tableStyle);
                            count++;
                        }
                    }
                    row = sheet1.createRow(count);
                    row.createCell(1).setCellValue("Total");
                    row.createCell(4).setCellValue(total);
                    sheet1.getRow(count).getCell(1).setCellStyle(tableStyle);
                    sheet1.getRow(count).getCell(4).setCellStyle(tableStyle);

                    sheet1.createRow(count + 2).createCell(1).setCellValue("The tax under the RCM process is 0");
                    row = sheet1.createRow(count + 4);
                    row.createCell(1).setCellValue("Account Name");
                    row.createCell(2).setCellValue(accountName);
                    row = sheet1.createRow(count + 5);
                    row.createCell(1).setCellValue("Account Number");
                    row.createCell(2).setCellValue(accountNo);
                    row = sheet1.createRow(count + 6);
                    row.createCell(1).setCellValue("Bank");
                    row.createCell(2).setCellValue(bankName);
                    row = sheet1.createRow(count + 7);
                    row.createCell(1).setCellValue("IFSC Code");
                    row.createCell(2).setCellValue(IFSCCode);
                    row = sheet1.createRow(count + 8);
                    row.createCell(1).setCellValue("PAN");
                    row.createCell(2).setCellValue(PANNo);
                    sheet1.getRow(count + 4).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(count + 4).getCell(2).setCellStyle(tableStyle);
                    sheet1.getRow(count + 5).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(count + 5).getCell(2).setCellStyle(tableStyle);
                    sheet1.getRow(count + 6).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(count + 6).getCell(2).setCellStyle(tableStyle);
                    sheet1.getRow(count + 7).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(count + 7).getCell(2).setCellStyle(tableStyle);
                    sheet1.getRow(count + 8).getCell(1).setCellStyle(tableStyleBold);
                    sheet1.getRow(count + 8).getCell(2).setCellStyle(tableStyle);

                    sheet1.createRow(count + 10).createCell(1).setCellValue("This is a system generated invoice hence not required any signature");
                    sheet1.getRow(count + 10).getCell(1).setCellStyle(cellStyleCenter);


                    wb.write(fileOut);
                    fileOut.close();

                    Workbook workbook = new Workbook(home + "/Desktop/ArtistTest/statement.xlsx");
                    Worksheet worksheet2 = workbook.getWorksheets().get(0);
                    worksheet2.autoFitColumns();
                    PdfSaveOptions options = new PdfSaveOptions();
                    options.setOnePagePerSheet(true);
                    workbook.save(home + "/Desktop/ArtistTest/OutputCorrectData/" + artistName + ".pdf", options);
                    Thread.sleep(1000);
                    break;
                }
            }
        }
        }
    }

}