package com.rakbank.excelparser.serviceImpl;


import com.rakbank.excelparser.model.Content;
import com.rakbank.excelparser.model.PatternPlaceHolders;
import com.rakbank.excelparser.model.WBSheet;
import com.rakbank.excelparser.service.WBSheetService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
@Slf4j
public class WBSheetServiceImpl implements WBSheetService {


    @Override
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
            String regex = "(?!\\(s\\))[^a-zA-Z0-9_,.& ]+[a-zA-Z0-9_ ]+[^a-zA-Z0-9_,.& ]+";
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

}
