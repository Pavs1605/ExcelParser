package com.rakbank.excelparser.serviceImpl;

import com.rakbank.excelparser.model.Content;
import com.rakbank.excelparser.model.PatternPlaceHolders;
import com.rakbank.excelparser.model.WBSheet;
import com.rakbank.excelparser.model.Spreadsheet;
import com.rakbank.excelparser.service.SpreadsheetService;
import com.rakbank.excelparser.service.WBSheetService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
@Slf4j
public class SpreadsheetServiceImpl implements SpreadsheetService {
    //    String filePath = "src/main/resources/static/SMSData.xlsx";
//
//    String outputFilePath = "src/main/resources/static/OutputTemplate.xlsx";
    @Value("${excel.file.path}")
    private String filePath;

    @Value("${excel.output.path}")
    private String outputFilePath;
    WBSheetService wbSheetService;

    @Autowired
    public SpreadsheetServiceImpl(WBSheetService wbSheetService) {
        this.wbSheetService = wbSheetService;
    }


    @Override
    public Spreadsheet extractValuesFromSpreadsheet() {
        Spreadsheet spreadsheet = new Spreadsheet();
        spreadsheet.setId(1);
        FileInputStream file = null;
        try {
            File fileObj = new File(filePath);
            file = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(file);

            if (fileObj == null || file == null) {
                log.error("extractValuesFromSpreadsheet() : Unable to find file : {}", filePath);
                return null;
            }
            //custom object
            spreadsheet.setSpreadSheetName(fileObj.getName());
            spreadsheet.setDefaultSheetName(workbook.getSheetAt(workbook.getActiveSheetIndex()).getSheetName());
            spreadsheet.setNoOfSheets(workbook.getNumberOfSheets());

            //setting data into the pojo's
            log.debug("extractValuesFromSpreadsheet(): Getting data from excel sheet");
            List<WBSheet> sheetsData = getDataFromAllSheets(workbook);
            spreadsheet.setSheets(sheetsData);

            //setting the data in output worksheet
            log.debug("extractValuesFromSpreadsheet(): Updating sheet with placeholders");
            updateSpreadSheet(sheetsData);


        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return spreadsheet;

    }

    private List<WBSheet> getDataFromAllSheets(Workbook workbook) {
        log.debug("getDataFromAllSheets(): Start getDataFromAllSheets");
        if (workbook == null)
            return null;

        List<WBSheet> sheetList = new ArrayList<>();
        int cnt = workbook.getNumberOfSheets();
        for (int i = 0; i < cnt; i++) {
            WBSheet sheetObj = new WBSheet();
            sheetObj.setId(i + 1);
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet != null)
                log.debug("getDataFromAllSheets(): Starting with sheet at index : {} sheetName :{}", i, (sheet != null ? sheet.getSheetName() : ""));

            //gets data from individual sheet
            sheetObj = wbSheetService.getData(sheet);
            sheetList.add(sheetObj);
        }
        return sheetList;

    }

    public void updateSpreadSheet(List<WBSheet> sheetsData) {
        log.debug("updateSpreadSheet(): Updating spreadsheet with extracted values");

        FileOutputStream outputStream = null;
        XSSFWorkbook workbook = null;

        try {
            log.debug("updateSpreadSheet(): Number of sheets to create: " + sheetsData.size());
            outputStream = new FileOutputStream(outputFilePath);
            workbook = new XSSFWorkbook();
            for (WBSheet singleSheet : sheetsData) {
                Sheet sheet = workbook.createSheet("Output-" + singleSheet.getName());
                sheet.setColumnWidth(0, 6000);
                sheet.setColumnWidth(1, 20000);
                sheet.setColumnWidth(2, 10000);
                sheet.setColumnWidth(3, 10000);

                //creating headers
                log.debug("updateSpreadSheet(): Creatng header rows ");
                Row header = sheet.createRow(0);

                CellStyle headerStyle = workbook.createCellStyle();
                headerStyle.setWrapText(true);

                XSSFFont font = workbook.createFont();
                font.setFontName("Arial");
                font.setFontHeightInPoints((short) 14);
                font.setBold(true);
                headerStyle.setFont(font);


                Cell headerCell = header.createCell(0);
                headerCell.setCellValue("Sno");
                headerCell.setCellStyle(headerStyle);

                headerCell = header.createCell(1);
                headerCell.setCellValue("Pattern");
                headerCell.setCellStyle(headerStyle);

                headerCell = header.createCell(2);
                headerCell.setCellValue("Event_RQ_Template");
                headerCell.setCellStyle(headerStyle);

                headerCell = header.createCell(3);
                headerCell.setCellValue("Event_id");
                headerCell.setCellStyle(headerStyle);


                //update with placeholder patterns
                log.debug("updateSpreadSheet(): Extracting placeholders and create rows with contents ");
                createRowsWithPlaceholders(singleSheet.getContentList(), sheet);
            }
            log.debug("updateSpreadSheet(): writing to workbook");
            workbook.write(outputStream);

        } catch (IOException e) {

        } finally {
            if (workbook != null) {

                try {
                    workbook.close();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }

            }

        }

    }

    private void createRowsWithPlaceholders(List<Content> contentList, Sheet sheet) {
        //this is needed to wrap text
        log.debug("createRowsWithPlaceholders(): Starting with createRows");

        Workbook wb = sheet.getWorkbook();
        CellStyle style = wb.createCellStyle();
        style.setWrapText(true);

        //getting placeholders
        int size = contentList.size();
        int rowNum = 1; //since 1st row is header row
        for (int i = 0; i < size; i++) {
            log.debug("createRowsWithPlaceholders(): creating rows & cell values");
            PatternPlaceHolders placeHolders = contentList.get(i).getPatternPlaceHolders();
            //if(!placeHolders.getEventRqTemplate().isEmpty()) {

            Row currRow = sheet.createRow(rowNum++); //since header row already exists
            Cell cell = currRow.createCell(0);
            cell.setCellValue(currRow.getRowNum());
            cell.setCellStyle(style);


            Cell patternCell = currRow.createCell(1);
            patternCell.setCellValue(placeHolders.getPattern());
            patternCell.setCellStyle(style);

            Cell eventRqCell = currRow.createCell(2);
            eventRqCell.setCellValue(placeHolders.getEventRqTemplate());
            eventRqCell.setCellStyle(style);

            Cell eventIdCell = currRow.createCell(3);
            eventIdCell.setCellValue(placeHolders.getEventId());
            eventIdCell.setCellStyle(style);
            //  }

        }
        log.debug("createRowsWithPlaceholders(): finished updating excel");
    }
}
