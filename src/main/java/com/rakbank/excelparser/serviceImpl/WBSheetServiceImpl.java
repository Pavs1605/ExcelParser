package com.rakbank.excelparser.serviceImpl;


import com.rakbank.excelparser.model.Content;
import com.rakbank.excelparser.model.PatternPlaceHolders;
import com.rakbank.excelparser.model.WBSheet;
import com.rakbank.excelparser.service.WBSheetService;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class WBSheetServiceImpl implements WBSheetService {

    String filePath = "src/main/resources/static/SMSData.xlsx";
    String outputPath = "src/main/resources/static/OutputTemplate.xlsx";

    @Override
    public WBSheet getData(Sheet sheet) {

        WBSheet wbSheet = new WBSheet();
        int rowCount = sheet.getLastRowNum() + 1;
        wbSheet.setName(sheet.getSheetName());
        wbSheet.setId(1);
        wbSheet.setRowCount(rowCount);

        // Assuming the first row contains column names
        Row headerRow = sheet.getRow(0);
        List<Content> contentList = new ArrayList<>();
        Map<Integer, String> columnIndexMap = new HashMap<>();
        for (Cell cell : headerRow) {
            String columnName = cell.getStringCellValue();
            int columnIndex = cell.getColumnIndex();
            columnIndexMap.put(columnIndex, columnName);
        }

        //assuming first is header row
        for (int i = 1; i < rowCount; i++) {
            Row row = sheet.getRow(i);

            Content content = new Content();
            content.setId(i);
            for (Cell cell : row) {
                int columnIndex = cell.getColumnIndex();
                String columnName = columnIndexMap.get(columnIndex);

                CellType type = cell.getCellType();

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

            PatternPlaceHolders placeHolders = extractValues(content);
            content.setPatternPlaceHolders(placeHolders);
            contentList.add(content);

        }

        wbSheet.setContentList(contentList);

        return wbSheet;
    }

    private PatternPlaceHolders extractValues(Content content) {
        String smsTemplate = content.getSmsTemplate();
        PatternPlaceHolders placeHolders = new PatternPlaceHolders();
        String newSmsTemplate = smsTemplate;
        String regexWordsNum = "[a-zA-Z0-9-,. ]";
        String regex = "[^a-zA-Z0-9_,. ]+[a-zA-Z_ ]+[^a-zA-Z0-9_,. ]+";
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
                System.out.println(matcherSpecial.group(0));
                String matchingStr = matcherSpecial.group(0);
                String placeHolderVal = removeSpecialCharacters(matchingStr);
                eventStr.append(placeHolderVal).append(":param").append(i++).append(";");
                newSmsTemplate = newSmsTemplate.replace(matchingStr, "(.*)");

            }

            placeHolders.setPattern(newSmsTemplate);
            placeHolders.setEventRqTemplate(eventStr.toString());
            placeHolders.setSno(content.getSno());
            placeHolders.setEventId(content.getEvent());

        }

        return placeHolders;
    }

    private String removeSpecialCharacters(String str) {
        String regexWords = "\\w+(_\\w+)+";//matches letters, numbers, underscores
        Pattern pattern = Pattern.compile(regexWords); // Match words with underscores
        Matcher matcher = pattern.matcher(str);
        String placeHolder = "";

        while (matcher.find()) {
            placeHolder = matcher.group();
            System.out.println(placeHolder);
            if (!placeHolder.isEmpty())
                return placeHolder;
        }

        return placeHolder;
    }

}
