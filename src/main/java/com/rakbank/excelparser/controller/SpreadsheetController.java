package com.rakbank.excelparser.controller;

import com.rakbank.excelparser.model.Spreadsheet;
import com.rakbank.excelparser.service.SpreadsheetService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/spreadsheets")
@Slf4j
public class SpreadsheetController {
    SpreadsheetService spreadsheetService;

    @Autowired
    public SpreadsheetController(SpreadsheetService spreadsheetService) {
        this.spreadsheetService = spreadsheetService;
    }
    @GetMapping
    public Spreadsheet extractValuesFromSpreadsheet() {
        log.debug("SpreadsheetController - In extract values from spreadsheet ");
        return spreadsheetService.extractValuesFromSpreadsheet();
    }

}
