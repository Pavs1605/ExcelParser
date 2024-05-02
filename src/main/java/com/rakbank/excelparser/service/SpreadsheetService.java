package com.rakbank.excelparser.service;

import com.rakbank.excelparser.model.Spreadsheet;
import org.springframework.stereotype.Service;

@Service
public interface SpreadsheetService {

    public Spreadsheet getDataFromSpreadsheet();
}
