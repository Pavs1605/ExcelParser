package com.rakbank.excelparser.serviceImpl;

import com.rakbank.excelparser.model.WBSheet;
import com.rakbank.excelparser.model.Spreadsheet;
import com.rakbank.excelparser.service.SpreadsheetService;
import com.rakbank.excelparser.service.WBSheetService;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
@Service
public class SpreadsheetServiceImpl implements SpreadsheetService {
    String filePath = "src/main/resources/static/SMSData.xlsx";
    String name = "SMSData.xlsx";
    WBSheetService wbSheetService;
@Autowired
    public SpreadsheetServiceImpl(WBSheetService wbSheetService) {
        this.wbSheetService = wbSheetService;
    }

    @Override
    public Spreadsheet getDataFromSpreadsheet() {
       Spreadsheet spreadsheet = new Spreadsheet();
       spreadsheet.setId(1);
        FileInputStream file = null;
        try {
            file = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(file);

          spreadsheet.setSpreadSheetName(name);
           spreadsheet.setDefaultSheetName(workbook.getSheetAt(workbook.getActiveSheetIndex()).getSheetName());
          spreadsheet.setNoOfSheets(workbook.getNumberOfSheets());
          spreadsheet.setSheets(getDataFromAllSheets(1));

          //saving
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return spreadsheet;

    }
    public List<WBSheet> getDataFromAllSheets(int spreadSheetId) {
        List<WBSheet> sheetList = new ArrayList<>();
        FileInputStream file = null;
        try {
            file = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(file);
            int cnt= workbook.getNumberOfSheets();
            for(int i=0;i<cnt;i++)
            {
                WBSheet sheetObj = new WBSheet();
                sheetObj.setId(i+1);
                Sheet sheet = workbook.getSheetAt(i);

                //gets data from sheet
                sheetObj =  wbSheetService.getData(sheet);
                sheetList.add(sheetObj);
            }

        }catch (IOException e) {
            throw new RuntimeException(e);
        }

        return sheetList;

    }

}
