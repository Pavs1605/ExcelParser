package com.rakbank.excelparser.util;


import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class ExcelParser {
    public static void main(String[] args) {
        ExcelParser excelParser = new ExcelParser();
        String filePath = "src/main/resources/static/SMSData.xlsx";
        String outputFilePath = "src/main/resources/static/OutputTemplate.xlsx";
        excelParser.extractFromSpreadSheet(filePath);
    }

    public void extractFromSpreadSheet(String filePath) {
        Spreadsheet spreadsheet = new Spreadsheet();
        spreadsheet.setId(1);
        FileInputStream file = null;
        try {
            File fileObj = new File(filePath);
            file = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(file);

            if (fileObj == null || file == null) {
                log.error("extractValuesFromSpreadsheet() : Unable to find file : {}", filePath);
                return;
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
            createOutputSpreadSheet(sheetsData, filePath);


        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        //  return spreadsheet;

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
            sheetObj = getData(sheet);
            sheetList.add(sheetObj);
        }
        return sheetList;

    }

    public void createOutputSpreadSheet(List<WBSheet> sheetsData, String filePath) {
        log.debug("createOutputSpreadSheet(): Updating spreadsheet with extracted values");

        FileOutputStream outputStream = null;

        XSSFWorkbook workbook = null;

        try {
            String directoryPath = new File(filePath).getParent();
            String tempFileName = "tempFile.xlsx";

            // Create a File object for the temporary file
            File tempFile = new File(directoryPath, "tempFile" + System.currentTimeMillis() + ".xlsx");
            log.debug("createOutputSpreadSheet(): Number of sheets to create: " + sheetsData.size());
            outputStream = new FileOutputStream(tempFile);
            workbook = new XSSFWorkbook();
            for (WBSheet singleSheet : sheetsData) {
                Sheet sheet = workbook.createSheet("Output-" + singleSheet.getName());
                sheet.setColumnWidth(0, 6000);
                sheet.setColumnWidth(1, 20000);
                sheet.setColumnWidth(2, 10000);
                sheet.setColumnWidth(3, 10000);

                //creating headers
                log.debug("createOutputSpreadSheet(): Creatng header rows ");
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
                log.debug("createOutputSpreadSheet(): Extracting placeholders and create rows with contents ");
                createRowsWithPlaceholders(singleSheet.getContentList(), sheet);
            }
            log.debug("createOutputSpreadSheet(): writing to workbook");
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
            if(!placeHolders.getEventRqTemplate().isEmpty()) {

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
              }

        }
        log.debug("createRowsWithPlaceholders(): finished updating excel");
    }

    public WBSheet getData(Sheet sheet) {
        log.debug("getData : extracting data from sheet");
        WBSheet wbSheet = new WBSheet();
        int rowCount = sheet.getLastRowNum() + 1;
        wbSheet.setName(sheet.getSheetName());
        wbSheet.setId(1);
        wbSheet.setRowCount(rowCount);

        // Assuming the first row contains column names
        Row headerRow = sheet.getRow(0);
        log.debug("getData : extracting columns form sheet");
        List<Content> contentList = new ArrayList<>();
        Map<Integer, String> columnIndexMap = new HashMap<>();
        for (Cell cell : headerRow) {
            String columnName = cell.getStringCellValue();
            int columnIndex = cell.getColumnIndex();
            columnIndexMap.put(columnIndex, columnName);
        }

        //assuming first is header row
        log.debug("getData : iterating throught rows to get content");
        for (int i = 1; i < rowCount; i++) {

            Row row = sheet.getRow(i);

            Content content = new Content();
            content.setId(i);
            for (Cell cell : row) {
                int columnIndex = cell.getColumnIndex();
                String columnName = columnIndexMap.get(columnIndex);

                CellType type = cell.getCellType();
                log.debug("getData : Getting data from rows and setting in content object");
                switch (columnName) {
                    case "Sno":
                        if (type.equals(CellType.NUMERIC)) {
                            content.setSno((int) cell.getNumericCellValue());
                        }
                        break;
                    case "Product":
                        if (type.equals(CellType.STRING)) {
                            content.setProduct(cell.getStringCellValue());
                        }
                        break;
                    case "Journey":
                        if (type.equals(CellType.STRING)) {
                            content.setJourney(cell.getStringCellValue());
                        }
                        break;
                    case "Event":
                        if (type.equals(CellType.STRING)) {
                            content.setEvent(cell.getStringCellValue());
                        }
                        break;
                    case "SMS Template":
                        if (type.equals(CellType.STRING)) {
                            content.setSmsTemplate(cell.getStringCellValue());
                        }
                        break;
                    default:
                        // Handle other columns if needed
                        break;
                }
            }
            log.debug("getData : Extracting placeholders from each row");
            PatternPlaceHolders placeHolders = extractValues(content);
            content.setPatternPlaceHolders(placeHolders);
            contentList.add(content);

            log.debug("getData : Setting placeholders in content object");

        }

        wbSheet.setContentList(contentList);

        return wbSheet;
    }


    public PatternPlaceHolders extractValues(Content content) {
        log.debug("extractValues : extracting values from rows");
        String smsTemplate = content.getSmsTemplate();
        PatternPlaceHolders placeHolders = new PatternPlaceHolders();
        String newSmsTemplate = smsTemplate;
        /*
        [a-zA-Z0-9-,. ] -> words having small or capital letters or numbers and which includes - , . and space
        (?!day\(s\)) -> this is to say exclude day(s)
        [^a-zA-Z0-9_,.& ]+ -> match for special characters, ^ -> negate op, looks for other characters apart from letters, small or caps, numbers,
                            or contains underscore, comma, fullstop, ampersand, + is atleast 1 match of special characters
         [a-zA-Z_ ]+ ->  can have 1 or more occurrence of leters, underscore
         */
        String regexWordsNum = "[a-zA-Z0-9-,. ]";
        String regex = "(?!\\(s\\))[^a-zA-Z0-9_,.& ]+[a-zA-Z_ ]+[^a-zA-Z0-9_,.& ]+";
        StringBuilder eventStr = new StringBuilder();
        int i = 0;

        Pattern patternWords = Pattern.compile(regexWordsNum);
        Pattern patternSpecialCharacters = Pattern.compile(regex);

        Matcher matcher = patternWords.matcher(smsTemplate);
        Matcher matcherSpecial = patternSpecialCharacters.matcher(smsTemplate);
        if (!matcher.find()) {
            System.out.println("No special characters found. Skipping pattern matching.");
            return placeHolders;
        } else {
            while (matcherSpecial.find()) {
                //  System.out.println(matcherSpecial.group(0));
                String matchingStr = matcherSpecial.group();
                String placeHolderVal = removeSpecialCharacters(matchingStr);
                eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
                newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

            }

            placeHolders.setPattern(newSmsTemplate);
            placeHolders.setEventRqTemplate(!eventStr.isEmpty() ? eventStr.substring(0, eventStr.length()-1) : "");
            placeHolders.setSno(content.getSno());
            placeHolders.setEventId(content.getEvent());

        }

        return placeHolders;
    }

    private String removeSpecialCharacters(String str) {
        //   String regexWords = "\\w+(_\\w+)+";//matches letters, numbers, underscores
        String regexWords = "[a-zA-Z0-9_]+";
        Pattern pattern = Pattern.compile(regexWords); // Match words with underscores
        Matcher matcher = pattern.matcher(str);
        String placeHolder = "";

        while (matcher.find()) {
            placeHolder = matcher.group();
            //  System.out.println(placeHolder);
            if (!placeHolder.isEmpty())
                return placeHolder;
        }

        return placeHolder;
    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public class Content {
        long id;
        int sno;
        String product;
        String journey;
        String event;
        String smsTemplate;
        PatternPlaceHolders patternPlaceHolders;
    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public class PatternPlaceHolders {
        int sno;
        String pattern;
        String eventRqTemplate;
        String eventId;
    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public class Spreadsheet {
        @Id
        @GeneratedValue(strategy = GenerationType.IDENTITY)
        long id;
        String spreadSheetName;
        String defaultSheetName;
        long noOfSheets;
        List<WBSheet> sheets;

    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public class WBSheet {
        long id;
        String name;
        long rowCount;
        long colCount;
        List<Content> contentList;
    }


}
