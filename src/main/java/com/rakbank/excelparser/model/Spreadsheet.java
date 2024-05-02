package com.rakbank.excelparser.model;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

@Data
@Slf4j
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
