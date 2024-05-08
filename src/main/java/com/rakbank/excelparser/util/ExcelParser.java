package com.rakbank.excelparser.util;


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
import java.util.concurrent.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


@Slf4j
public class ExcelParser {
    ExecutorService executorService = Executors.newFixedThreadPool(10);
    public static final String XLSX = ".xlsx";
    public static final String EVENT_ID_COL_NAME = "EVENT_ID";

    public static final String SNO_COL_NAME = "SNO";
    public static final String SMS_TEMPLATE_COL_NAME = "SMS TEMPLATE";
    public static final String PATTERN_COL_NAME = "PATTERN";
    public static final String EVENT_RQ_TEMPLATE_COL_NAME = "EVENT_RQ_TEMPLATE";
    public static final String ORIGINAL_STR_COL_NAME = "Original String";

    public static final String OUTPUT_FILE_NAME = "tempFile";
    public static final int SMS_TEMPLATE_COL_NO = 3; //defaulting 5th column for sms template col
    public static final int EVENT_ID_COL_NO = 0; //defaulting 5th column for sms template col

    public static final boolean combineTabs = true;

    public static Pattern[] patterns = {};

    public static final String[] includeStrings = { "@"}; //Say for <@CardNumber> -> the param should be @CardNumber, then add character here.
    public static final String[] excludeStrings = { "/"};
    public static  String commonRegex = "[A-Za-z0-9_ ]+"; // Extracts letters numbers,_, and space, strings like prospect_id or prospect id is extracted

    static {
        initializeRegex();
    }

    public static void initializeRegex() {
        StringBuilder regexBuilder = new StringBuilder();

        // Append the common regex part
        regexBuilder.append(commonRegex);

        // Append the include strings
        for (String include : includeStrings) {
            regexBuilder.append("|").append(Pattern.quote(include));
        }

        // Append the exclude strings as negative lookahead assertions
        for (String exclude : excludeStrings) {
            regexBuilder.append("(?!.*").append(Pattern.quote(exclude)).append(")");
        }

        String commonRegex = regexBuilder.toString();


       patterns = new Pattern[]{
               Pattern.compile("([~@$%*+=\\-&\\#!\\^\\`\\?\\:\\|]+)\\s*" + commonRegex + "\\1"), // Special characters
               Pattern.compile("\\[+" + commonRegex + "\\]+"), // Square brackets
               Pattern.compile("\\(+" + commonRegex + "\\)+"), // Parentheses
               Pattern.compile("\\{+" + commonRegex + "\\}+"), // Curly braces
               Pattern.compile("\\<+" + commonRegex + "\\>+") // Angular brackets
       };
    }

    public static void main(String[] args) {
        ExcelParser excelParser = new ExcelParser();
        String filePath = "src/main/resources/static/SMS_Email_Templates_IVR_2.xlsx";

        long startTime = System.currentTimeMillis();
        log.debug("Start time : " + startTime);

        //core logic
        excelParser.extractFromSpreadSheet(filePath);
        //shutting down executor service
        excelParser.shutdownExecutorService();
        log.debug("Finished writing in spreadsheet");

        //calculating time
        long endTime = System.currentTimeMillis();
        log.debug("end time : " + endTime);

        long totalTime = endTime - startTime;
        log.debug("Total Time taken in seconds::" + (double) (totalTime) / (double) (1000) + ", In minutes ::" + (double) (totalTime) / (double) (1000 * 60));

    }

    public void extractFromSpreadSheet(String filePath) {

        try {
            File fileObj = new File(filePath);
            FileInputStream file = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(file);

            log.debug("Inside extract from spreadsheet, reading from Excel sheet");
            if (fileObj == null || file == null) {
                log.error("extractValuesFromSpreadsheet() : Unable to find file : {}", filePath);
                return;
            }

            //custom object
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.setSpreadSheetName(fileObj.getName());
            spreadsheet.setNoOfSheets(workbook.getNumberOfSheets());
            log.debug("Spreadsheet object :: " + spreadsheet);

            //setting data into the pojo's from sheets
            log.debug("extractValuesFromSpreadsheet(): Getting data from excel sheet");
            List<WBSheet> sheetsData = getDataFromAllTabs(workbook);
            if (sheetsData == null) {
                log.error("Sheets data is null");
                log.debug("Parser failed, unable to get data from sheets");
                return;
            }
            spreadsheet.setSheets(sheetsData);

            //setting the data in output worksheet
            log.debug("extractValuesFromSpreadsheet(): Updating sheet with placeholders");
            if (combineTabs) {
                createOutputSpreadsheet(sheetsData, filePath);
            } else {
                createSingleOutputSpreadsheet(sheetsData, filePath);
            }

           log.debug("Exiting writing in spreadsheet");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        //  return spreadsheet;

    }

    private List<WBSheet> getDataFromAllTabs(Workbook workbook) {

        log.debug("getDataFromAllSheets(): Start getDataFromAllSheets");
        log.debug("Getting data from all sheets");
        List<WBSheet> sheetList = new ArrayList<>();
        try {
            if (workbook == null)
                return null;

            int cnt = workbook.getNumberOfSheets();

            List<CompletableFuture<WBSheet>> futures = new ArrayList<>();
            for (int i = 0; i < cnt; i++) {
                int finalI = i;
                CompletableFuture<WBSheet> future = CompletableFuture.supplyAsync(() -> {
                    Sheet sheet = workbook.getSheetAt(finalI);
                    if (sheet == null) {
                        log.debug("sheet is null at index ::" + finalI + ";exiting code");
                        log.debug("getDataFromAllSheets(): Starting with sheet at index : {} sheetName :{}", finalI, "");
                        return null;
                    }
                    // gets data from individual sheet
                    return getDataFromSingleSheet(sheet);
                }, executorService);

                futures.add(future);
            }

            CompletableFuture<Void> allFutures = CompletableFuture.allOf(futures.toArray(new CompletableFuture[0]));
            allFutures.get(); // Wait for all tasks to complete

            log.debug("Streaming all futures");
            futures.stream()
                    .map(CompletableFuture::join) // Get the result of each CompletableFuture
                    .filter(Objects::nonNull) // Filter out any null results
                    .forEach(sheetList::add); // Add valid results to sheetList

            System.out.println("sheet list size" + sheetList.size());

        } catch (Exception e) {
            log.error("Error waiting for all tasks to complete: {}", e.getMessage());
        }
        return sheetList;


    }

    private void createOutputSpreadsheet(List<WBSheet> sheetsData, String filePath) {

        if (sheetsData.isEmpty()) {
            System.out.println("No sheets data found");
            return;
        }

        log.debug("createOutputSpreadSheet(): Updating spreadsheet with extracted values");
        XSSFWorkbook workbook = null;
        try {
            String directoryPath = new File(filePath).getParent();
            String tempFileName = OUTPUT_FILE_NAME + System.currentTimeMillis() + XLSX;
            File tempFile = new File(directoryPath, tempFileName);

            FileOutputStream outputStream = new FileOutputStream(tempFile);
            workbook = new XSSFWorkbook();

            for (WBSheet singleSheet : sheetsData) {

               Sheet sheet = workbook.createSheet("Output-" + singleSheet.getName());

                //creating headers
                creatingHeaderRowInOutputSheet(sheet, workbook);

                //update with placeholder patterns
                log.debug("createOutputSpreadSheet(): Extracting placeholders and create rows with contents ");
                createRowsInSheetWithDynamicValues(singleSheet.getRowContentList(), sheet);

            }
            log.debug("createOutputSpreadSheet(): writing to workbook");
            workbook.write(outputStream);

        } catch (IOException e) {
           log.error(e.getMessage());

        } finally {
            if (workbook != null) {

                try {
                    workbook.close();
                } catch (IOException e) {
                    log.error(e.getMessage());

                }

            }

        }

    }

    private void createSingleOutputSpreadsheet(List<WBSheet> sheetsData, String filePath) {
        log.debug("createOutputSpreadSheet(): Updating spreadsheet with extracted values");
        if (sheetsData.isEmpty()) {
           log.error("No sheets data found");
            return;
        }

        XSSFWorkbook workbook = null;

        try {
            String directoryPath = new File(filePath).getParent();
            String tempFileName = OUTPUT_FILE_NAME + System.currentTimeMillis() + XLSX;
            File tempFile = new File(directoryPath, tempFileName);
            FileOutputStream outputStream = new FileOutputStream(tempFile);

            workbook = new XSSFWorkbook();

            Sheet sheet = workbook.createSheet("Output-" + sheetsData.get(0).getName());
            //creating headers
            creatingHeaderRowInOutputSheet(sheet, workbook);

            for (WBSheet singleSheet : sheetsData) {
                //update with placeholder patterns
                log.debug("createOutputSpreadSheet(): Extracting placeholders and create rows with contents ");
                createRowsInSheetWithDynamicValues(singleSheet.getRowContentList(), sheet);
            }
            log.debug("createOutputSpreadSheet(): writing to workbook");
            workbook.write(outputStream);

        } catch (IOException e) {
            log.error(e.getMessage());

        } finally {
            if (workbook != null) {

                try {
                    workbook.close();
                } catch (IOException e) {
                    log.error(e.getMessage());

                }

            }

        }

    }
    private void creatingHeaderRowInOutputSheet(Sheet sheet, XSSFWorkbook workbook) {
        log.debug("createOutputSpreadSheet(): Creating header rows ");
        Row header = sheet.createRow(0);

        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 20000);
        sheet.setColumnWidth(2, 10000);
        sheet.setColumnWidth(3, 10000);
        sheet.setColumnWidth(4, 20000);


        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setWrapText(true);

        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        headerStyle.setFont(font);


        Cell headerCell = header.createCell(0);
        headerCell.setCellValue(SNO_COL_NAME);
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue(PATTERN_COL_NAME);
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(2);
        headerCell.setCellValue(EVENT_RQ_TEMPLATE_COL_NAME);
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(3);
        headerCell.setCellValue(EVENT_ID_COL_NAME);
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(4);
        headerCell.setCellValue(ORIGINAL_STR_COL_NAME);
        headerCell.setCellStyle(headerStyle);
    }


    private void createRowsInSheetWithDynamicValues(List<RowContent> rowContentList, Sheet sheet) {
        log.debug("createRowsWithPlaceholders(): Starting with createRows");

        Workbook wb = sheet.getWorkbook();
        CellStyle style = wb.createCellStyle();
        style.setWrapText(true); //to wrap text

        int rowNum = 1; //since 1st row is header row
        if (rowContentList == null || rowContentList.isEmpty()) {
            System.out.println("Unable to create as content list is empty");
            return;
        }
        for (RowContent rowContent : rowContentList) {
            log.debug("createRowsWithPlaceholders(): creating rows & cell values");
            DynamicRowContent placeHolders = rowContent.getDynamicRowContent();
            if (placeHolders == null)
                continue;

            Row currRow = sheet.createRow(rowNum++); //since header row already exists, starting from 1st row
            Cell cell = currRow.createCell(0);
            cell.setCellValue(currRow.getRowNum());
            cell.setCellStyle(style);

            Cell patternCell = currRow.createCell(1);
            String pat = placeHolders.getPattern();
            patternCell.setCellValue(pat.isEmpty() ? "" : pat);
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

        }
        log.debug("createRowsWithPlaceholders(): finished updating excel");
    }

    private WBSheet getDataFromSingleSheet(Sheet sheet) {
        log.debug("getData : extracting data from sheet");
        if (sheet == null){
            log.debug("getDataFromSingleSheet : Sheet obj does not have data");
            return null;
        }

        int smsTemplateColumnNo = SMS_TEMPLATE_COL_NO;
        int eventIdColNo = EVENT_ID_COL_NO;
        boolean smsTemplateColumnExists = false;
        boolean eventIdColumnExists = false;
        int headerRowNo = 1;
        //checking if SMSTemplate column is found


        Map<Integer, String> columnIndexMap = new HashMap<>();

        int rowCount = sheet.getLastRowNum() + 1;
        //creating WBSheet object
        WBSheet wbSheet = new WBSheet();
        wbSheet.setName(sheet.getSheetName());

        log.debug("getData : extracting columns from sheet");


        for (int i = 0; i < rowCount; i++) {

            Row row = sheet.getRow(i);
            for (Cell cell : row) {
                String columnName = cell.getStringCellValue().trim();
                int columnIndex = cell.getColumnIndex();
                columnIndexMap.put(columnIndex, columnName);
                System.out.println("Column name ::" + columnName + ", Column index ::" + columnIndex);


                if (columnName.equalsIgnoreCase(SMS_TEMPLATE_COL_NAME)) {
                    smsTemplateColumnExists = true;
                    smsTemplateColumnNo = columnIndex;
                    headerRowNo = i;
                    log.debug("Sms template col found on :: " + smsTemplateColumnNo);
                    log.debug("Header row number :: " + headerRowNo);

                }
                if (columnName.equalsIgnoreCase(EVENT_ID_COL_NAME)) {
                    eventIdColumnExists = true;
                    eventIdColNo = columnIndex;
                    log.debug("Sms template col found on :: " + smsTemplateColumnNo);
                    log.debug("Header row number :: " + headerRowNo);

                }


            }
        }

        if (!smsTemplateColumnExists || !eventIdColumnExists) {
            System.out.println("Sms template col or event id column does not exist");
            return null;
        }


        //ignoring header row for parsing data
        List<RowContent> rowContentList = parseRowAndExtractDynamicValue(sheet, headerRowNo, smsTemplateColumnNo, eventIdColNo, columnIndexMap);
        wbSheet.setRowContentList(rowContentList);

        return wbSheet;
    }

    private List<RowContent> parseRowAndExtractDynamicValue(Sheet sheet, int headerRowNo, int smsTemplateColumnNo, int eventIdColNo, Map<Integer, String> columnIndexMap) {
        log.debug("getData : iterating through rows to get content");
        List<RowContent> rowContentList = new ArrayList<>();
        int rowCount = sheet.getLastRowNum() + 1;
        for (int i = headerRowNo + 1; i < rowCount; i++) {

            Row row = sheet.getRow(i);

            boolean validationPassed = validateRow(row, smsTemplateColumnNo, eventIdColNo);

            if (!validationPassed) {
                System.out.println("Validation failed, not a valid row, checking next row");
                continue;
            }
            RowContent rowContent = new RowContent();
            for (Cell cell : row) {
                int columnIndex = cell.getColumnIndex();
                String columnName = columnIndexMap.get(columnIndex);

                CellType type = cell.getCellType();
                log.debug("getData : Getting data from rows and setting in content object");
                switch (columnName) {
                    case SNO_COL_NAME:
                        if (type.equals(CellType.NUMERIC)) {
                            rowContent.setSno((int)cell.getNumericCellValue());
                        }
                        break;
                    case EVENT_ID_COL_NAME:
                        if (type.equals(CellType.STRING)) {
                            rowContent.setEvent(cell.getStringCellValue());
                        }
                        break;
                    case SMS_TEMPLATE_COL_NAME:
                        if (type.equals(CellType.STRING)) {
                            rowContent.setSmsTemplate(cell.getStringCellValue());
                        }
                        break;
                    default:
                        // Handle other columns if needed
                        break;
                }
            }

            if (rowContent.getSmsTemplate() == null || rowContent.getSmsTemplate().isEmpty() ||
                    rowContent.getEvent() == null || rowContent.getEvent().isEmpty()){
                log.error("No sms template column");
            } else {
                log.debug("header row content ::" + rowContent);
                log.debug("getData : Extracting placeholders from each row");
                DynamicRowContent placeHolders = extractDynamicValuesFromSmsTemplate(rowContent);
                rowContent.setDynamicRowContent(placeHolders);
                rowContentList.add(rowContent);
                log.debug("getData : Setting placeholders in content object");

            }

        }
        return rowContentList;
    }

    private boolean validateRow(Row row, int smsTemplateColNo, int eventIdNo) {
        if (row == null) {
            System.out.println("row is empty, validation failed");
            return false;
        }

        //checking if sms template value is not null
        Cell cellAtSmsTemp = row.getCell(smsTemplateColNo);
        Cell cellATEventId = row.getCell(eventIdNo);
        if (cellAtSmsTemp == null || cellAtSmsTemp.getStringCellValue() == null || cellATEventId.getStringCellValue() == null || cellATEventId.getStringCellValue().isEmpty()) {
            System.out.println("Sms template does not have content, not a valid row");
            return false;
        }

        return true;
    }

    public DynamicRowContent extractDynamicValuesFromSmsTemplate(RowContent rowContent) {
        log.debug("extractValues: extracting values from rows");
        System.out.println("extractDynamicValuesFromSmsTemplate");

        String smsTemplate = rowContent.getSmsTemplate();
        DynamicRowContent placeHolders = new DynamicRowContent();
        String newSmsTemplate = smsTemplate;

        StringBuilder eventStr = new StringBuilder();
        int i = 0;

        for (Pattern pattern : patterns) {
            Matcher matcher = pattern.matcher(smsTemplate);
            System.out.println("pattern used :: " + pattern);

            // Iterate through matcher to replace placeholders
            while (matcher.find()) {
                String matchingStr = matcher.group();
                System.out.println("found match :" + matchingStr);

                String placeHolderVal = removeSpecialCharactersFromDynamicValues(matchingStr);
                eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
                newSmsTemplate = newSmsTemplate.replaceAll(Pattern.quote(matchingStr), "(.*)");
            }
        }

        placeHolders.setPattern(newSmsTemplate);
        placeHolders.setEventRqTemplate(!eventStr.isEmpty() ? eventStr.substring(0, eventStr.length() - 1) : "");
        placeHolders.setSno(rowContent.getSno());
        placeHolders.setEventId(rowContent.getEvent());
        placeHolders.setOriginalString(rowContent.getSmsTemplate());

        System.out.println("placeholder object :" + placeHolders);

        return placeHolders;
    }


    private DynamicRowContent extractDynamicValuesFromSmsTemplateOld(RowContent rowContent) {
        log.debug("extractValues : extracting values from rows");
        System.out.println("extractDynamicValuesFromSmsTemplate");

        String smsTemplate = rowContent.getSmsTemplate();
        DynamicRowContent placeHolders = new DynamicRowContent();
        String newSmsTemplate = smsTemplate;

        String regexWordsNum = "[a-zA-Z0-9-,. ]";
        String regex = "\\b([~@$%*+=\\-&#!]+)([A-Za-z0-9_ ]+)\\1\\b"; //working one
        String regexForSquareBraces = "\\[+([A-Za-z0-9_ ]+)\\]+";
        String regexForParenthesis = "\\(+([A-Za-z0-9_ ]+)\\)+";
        String regexForCurlyBraces = "\\{+([A-Za-z0-9_ ]+)\\}+";
        String regexForAngularBraces = "\\<+([A-Za-z0-9_ ]+)\\>+";


        StringBuilder eventStr = new StringBuilder();
        int i = 0;

        Pattern patternWords = Pattern.compile(regexWordsNum);
        Pattern patternSpecialCharacters = Pattern.compile(regex);
        Pattern patternSquareBraces = Pattern.compile(regexForSquareBraces);
        Pattern patternParenthesis = Pattern.compile(regexForParenthesis);
        Pattern patternCurlyBraces = Pattern.compile(regexForCurlyBraces);
        Pattern patternAngularBraces = Pattern.compile(regexForAngularBraces);


        Matcher matcher = patternWords.matcher(smsTemplate);
        Matcher matcherSpecial = patternSpecialCharacters.matcher(smsTemplate);
        Matcher matcherSquareBraces = patternSquareBraces.matcher(smsTemplate);
        Matcher matcherParenthesis = patternParenthesis.matcher(smsTemplate);
        Matcher matcherCurlyBraces = patternCurlyBraces.matcher(smsTemplate);
        Matcher matcherAngularBraces = patternAngularBraces.matcher(smsTemplate);

        if (!matcher.find()) {
            System.out.println("No special characters found. Skipping pattern matching.");
            return placeHolders;
        } else if (matcherSpecial.find()) {
            while (matcherSpecial.find()) {
                //  System.out.println(matcherSpecial.group(0));
                String matchingStr = matcherSpecial.group();
                System.out.println("found match :" + matchingStr);

                String placeHolderVal = removeSpecialCharactersFromDynamicValues(matchingStr);
                eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
                newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

            }
        } else if (matcherSquareBraces.find()) {
            //  System.out.println(matcherSpecial.group(0));
            String matchingStr = matcherSpecial.group();
            System.out.println("found match :" + matchingStr);

            String placeHolderVal = removeSpecialCharactersFromDynamicValues(matchingStr);
            eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
            newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

        } else if (matcherParenthesis.find()) {
            //  System.out.println(matcherSpecial.group(0));
            String matchingStr = matcherSpecial.group();
            System.out.println("found match :" + matchingStr);

            String placeHolderVal = removeSpecialCharactersFromDynamicValues(matchingStr);
            eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
            newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

        } else if (matcherCurlyBraces.find()) {
            //  System.out.println(matcherSpecial.group(0));
            String matchingStr = matcherSpecial.group();
            System.out.println("found match :" + matchingStr);

            String placeHolderVal = removeSpecialCharactersFromDynamicValues(matchingStr);
            eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
            newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

        } else if (matcherAngularBraces.find()) {
            //  System.out.println(matcherSpecial.group(0));
            String matchingStr = matcherSpecial.group();
            System.out.println("found match :" + matchingStr);

            String placeHolderVal = removeSpecialCharactersFromDynamicValues(matchingStr);
            eventStr.append(placeHolderVal).append(":param").append(i++).append(",");
            newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

        }

        placeHolders.setPattern(newSmsTemplate);
        placeHolders.setEventRqTemplate(!eventStr.isEmpty() ? eventStr.substring(0, eventStr.length() - 1) : "");
        placeHolders.setSno(rowContent.getSno());
        placeHolders.setEventId(rowContent.getEvent());
        placeHolders.setOriginalString(rowContent.getSmsTemplate());

        System.out.println("placeholder object :" + placeHolders);

        return placeHolders;


    }

    private String removeSpecialCharactersFromDynamicValues(String str) {
        System.out.println("removeSpecialCharactersFromDynamicValues from :: " + str);

        Pattern pattern = Pattern.compile(commonRegex); // Match words with underscores
        Matcher matcher = pattern.matcher(str);
        String placeHolder = "";

        while (matcher.find()) {
            placeHolder = matcher.group();
            if (!placeHolder.isEmpty())
                return placeHolder;
        }

        System.out.println("removed special characters final value :: " + placeHolder);

        return placeHolder;
    }




    private void shutdownExecutorService() {
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
    public class RowContent {
        int sno;
        String event;
        String smsTemplate;
        DynamicRowContent dynamicRowContent;
    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public class DynamicRowContent {
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
            long id;
            String spreadSheetName;
            long noOfSheets;
            List<WBSheet> sheets;

        }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public class WBSheet {
        String name;
        List<RowContent> rowContentList;
    }


}
