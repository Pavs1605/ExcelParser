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
public class WBSheet {
    long id;
    String name;
    long rowCount;
    long colCount;
    List<Content> contentList;
}
