package com.rakbank.excelparser.service;


import com.rakbank.excelparser.model.WBSheet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.util.List;
@Service
public interface WBSheetService {

    public WBSheet getData(Sheet sheet);

    }
