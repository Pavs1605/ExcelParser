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
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class ExcelParser {
    ExecutorService executorService = Executors.newFixedThreadPool(10);
    public static void main(String[] args) {
        ExcelParser excelParser = new ExcelParser();
        String filePath = "src/main/resources/static/SMSData.xlsx";

       // String outputFilePath = "src/main/resources/static/OutputTemplate.xlsx";
        long startTime = System.currentTimeMillis();
        //System.out.println("Start time : " + startTime);

        //core logic
        excelParser.extractFromSpreadSheet(filePath);
        excelParser.shutdownExecutorService(); // Shutdown the ExecutorService
        //System.out.println("Finished writing in spreadsheet");

        //calculating time
        long endTime = System.currentTimeMillis();
        //System.out.println("end time : " + endTime);

        long totalTime = endTime-startTime;
        System.out.println("Total Time taken in seconds::" + (double)(totalTime)/ (double)(1000) + ", In minutes ::" +  (double)(totalTime)/ (double)(1000*60));


    }

    public void extractFromSpreadSheet(String filePath) {
        Spreadsheet spreadsheet = new Spreadsheet();
        spreadsheet.setId(1);

        try {
            File fileObj = new File(filePath);
            FileInputStream  file = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(file);

            //System.out.println("Inside extract from spreadsheet, reading from excel sheet");
            if (fileObj == null || file == null) {
                log.error("extractValuesFromSpreadsheet() : Unable to find file : {}", filePath);
                //System.out.println("Unable to find file for extracting exiting");
                return;
            }
            //custom object
            spreadsheet.setSpreadSheetName(fileObj.getName());
            spreadsheet.setDefaultSheetName(workbook.getSheetAt(workbook.getActiveSheetIndex()).getSheetName());
            spreadsheet.setNoOfSheets(workbook.getNumberOfSheets());
            //System.out.println("Spreadsheet object :: " + spreadsheet);


            //setting data into the pojo's
            log.debug("extractValuesFromSpreadsheet(): Getting data from excel sheet");
            List<WBSheet> sheetsData = getDataFromAllSheets(workbook);
            if(sheetsData == null){
                log.error("Sheets data is null");
                //System.out.println("Parser failed, unable to get data from sheets");
                return;
            }
            //System.out.println("Spreadsheet object is set ");
            spreadsheet.setSheets(sheetsData);

            //setting the data in output worksheet
            log.debug("extractValuesFromSpreadsheet(): Updating sheet with placeholders");
            //System.out.println("Creating new spreadsheet");
            createOutputInNewSpreadSheet(sheetsData, filePath);

            //System.out.println("Exiting writing in spreadsheet");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        //  return spreadsheet;

    }

    private List<WBSheet> getDataFromAllSheets(Workbook workbook) {

        log.debug("getDataFromAllSheets(): Start getDataFromAllSheets");
        //System.out.println("Getting data from all sheets");
        List<WBSheet> sheetList = new ArrayList<>();
        if (workbook == null)
            return null;

        int cnt = workbook.getNumberOfSheets();


        List<CompletableFuture<WBSheet>> futures = new ArrayList<>();
        for (int i = 0; i < cnt; i++) {
            int finalI = i;
            CompletableFuture<WBSheet> future = CompletableFuture.supplyAsync(() -> {
                WBSheet sheetObj = new WBSheet();
                sheetObj.setId(finalI + 1);
                Sheet sheet = workbook.getSheetAt(finalI);
                if (sheet == null) {
                    //System.out.println("sheet is null at index ::" + finalI + "exiting code");
                    log.debug("getDataFromAllSheets(): Starting with sheet at index : {} sheetName :{}", finalI, "");
                    return null;
                }
                // gets data from individual sheet
                return getDataFromSheet(sheet);
            }, executorService);

            futures.add(future);
        }

        CompletableFuture<Void> allFutures = CompletableFuture.allOf(futures.toArray(new CompletableFuture[0]));

        try {
            allFutures.get(); // Wait for all tasks to complete
        } catch (Exception e) {
            log.error("Error waiting for all tasks to complete: {}", e.getMessage());
        }

        //System.out.println("Streaming all futures");
        futures.stream()
                .map(CompletableFuture::join) // Get the result of each CompletableFuture
                .filter(Objects::nonNull) // Filter out any null results
                .forEach(sheetList::add); // Add valid results to sheetList

        //System.out.println("sheet list size" + sheetList.size());
        return sheetList;

    }

    public void createOutputInNewSpreadSheet(List<WBSheet> sheetsData, String filePath) {
        log.debug("createOutputSpreadSheet(): Updating spreadsheet with extracted values");


        XSSFWorkbook workbook = null;

        try {
            String directoryPath = new File(filePath).getParent();
            String tempFileName = "tempFile" + System.currentTimeMillis() + ".xlsx";
            // Create a File object for the temporary file
            File tempFile = new File(directoryPath, tempFileName);
            log.debug("createOutputSpreadSheet(): Number of sheets to create: " + sheetsData.size());
            FileOutputStream  outputStream = new FileOutputStream(tempFile);
            workbook = new XSSFWorkbook();

            for (WBSheet singleSheet : sheetsData) {
                Sheet sheet = workbook.createSheet("Output-" + singleSheet.getName());
                sheet.setColumnWidth(0, 6000);
                sheet.setColumnWidth(1, 20000);
                sheet.setColumnWidth(2, 10000);
                sheet.setColumnWidth(3, 10000);
                sheet.setColumnWidth(4, 20000);

                //creating headers
                log.debug("createOutputSpreadSheet(): Creating header rows ");
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

                headerCell = header.createCell(4);
                headerCell.setCellValue("Original SMS template");
                headerCell.setCellStyle(headerStyle);


                //update with placeholder patterns
                log.debug("createOutputSpreadSheet(): Extracting placeholders and create rows with contents ");
                createRowsWithPlaceholders(singleSheet.getContentList(), sheet);
            }
            log.debug("createOutputSpreadSheet(): writing to workbook");
            workbook.write(outputStream);

        } catch (IOException e) {
            e.printStackTrace();

        } finally {
            if (workbook != null) {

                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();

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

        int rowNum = 1; //since 1st row is header row
        for (Content content : contentList) {
            log.debug("createRowsWithPlaceholders(): creating rows & cell values");
            PatternPlaceHolders placeHolders = content.getPatternPlaceHolders();
         //   if (!placeHolders.getEventRqTemplate().isEmpty()) {

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

                Cell originalStrCell = currRow.createCell(4);
            originalStrCell.setCellValue(placeHolders.getOriginalString());
            originalStrCell.setCellStyle(style);
          //  }

        }
        log.debug("createRowsWithPlaceholders(): finished updating excel");
    }

    public WBSheet getDataFromSheet(Sheet sheet) {
        log.debug("getData : extracting data from sheet");
        //System.out.println("Extracting data from sheet");
        WBSheet wbSheet = new WBSheet();
        int rowCount = sheet.getLastRowNum() + 1;
        wbSheet.setName(sheet.getSheetName());
        wbSheet.setId(1);
        wbSheet.setRowCount(rowCount);

        // Assuming the first row contains column names
        Row headerRow = sheet.getRow(0);
        //System.out.println("Setting header row");
        log.debug("getData : extracting columns form sheet");
        List<Content> contentList = new ArrayList<>();
        Map<Integer, String> columnIndexMap = new HashMap<>();
        for (Cell cell : headerRow) {
            String columnName = cell.getStringCellValue();
            int columnIndex = cell.getColumnIndex();
            columnIndexMap.put(columnIndex, columnName);
        }

        //assuming first is header row
        log.debug("getData : iterating through rows to get content");
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
                    case "event":
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

            //System.out.println("header row content ::" + content );
            log.debug("getData : Extracting placeholders from each row");
            PatternPlaceHolders placeHolders = extractDynamicValuesFromSmsTemplate(content);
            content.setPatternPlaceHolders(placeHolders);
            contentList.add(content);

            log.debug("getData : Setting placeholders in content object");

        }

        wbSheet.setContentList(contentList);

        return wbSheet;
    }


    public PatternPlaceHolders extractDynamicValuesFromSmsTemplate(Content content) {
        log.debug("extractValues : extracting values from rows");
        //System.out.println("extractDynamicValuesFromSmsTemplate" );

        String smsTemplate = content.getSmsTemplate();
        PatternPlaceHolders placeHolders = new PatternPlaceHolders();
        String newSmsTemplate = smsTemplate;
        /*
        [a-zA-Z0-9-,. ] -> words having small or capital letters or numbers and which includes - or , or . and space
        (?!day\(s\)) -> this is to say exclude day(s)
        [^a-zA-Z0-9_,.& ]+ -> match for special characters, ^ -> negate op, looks for other characters apart from letters, small or caps, numbers,
                            or contains underscore, comma, full stop, ampersand, + is at least 1 match of special characters
         [a-zA-Z_ ]+ ->  can have 1 or more occurrence of letters, underscore
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
            //System.out.println("No special characters found. Skipping pattern matching.");
            return placeHolders;
        } else {
            while (matcherSpecial.find()) {
                //  //System.out.println(matcherSpecial.group(0));
                String matchingStr = matcherSpecial.group();
                //System.out.println("found match :" + matchingStr );

                String placeHolderVal = removeSpecialCharactersFromDynamicValues(matchingStr);
                eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
                newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

            }

            placeHolders.setPattern(newSmsTemplate);
            placeHolders.setEventRqTemplate(!eventStr.isEmpty() ? eventStr.substring(0, eventStr.length()-1) : "");
            placeHolders.setSno(content.getSno());
            placeHolders.setEventId(content.getEvent());
            placeHolders.setOriginalString(content.getSmsTemplate());

            //System.out.println("placeholder object :" + placeHolders );


        }

        return placeHolders;
    }

    private String removeSpecialCharactersFromDynamicValues(String str) {
        //System.out.println("removeSpecialCharactersFromDynamicValues from :: " + str );

        //   String regexWords = "\\w+(_\\w+)+";//matches letters, numbers, underscores
        String regexWords = "[a-zA-Z0-9_]+";
        Pattern pattern = Pattern.compile(regexWords); // Match words with underscores
        Matcher matcher = pattern.matcher(str);
        String placeHolder = "";

        while (matcher.find()) {
            placeHolder = matcher.group();
            //  //System.out.println(placeHolder);
            if (!placeHolder.isEmpty())
                return placeHolder;
        }

        //System.out.println("removed special characters final value :: " + placeHolder );

        return placeHolder;
    }

    public void shutdownExecutorService() {
        executorService.shutdown();
        try {

            if (!executorService.awaitTermination(10000, TimeUnit.SECONDS)) {
                executorService.shutdownNow();
            }
        } catch (InterruptedException e) {
            executorService.shutdownNow();
            Thread.currentThread().interrupt();
        }
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
        String originalString;
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
