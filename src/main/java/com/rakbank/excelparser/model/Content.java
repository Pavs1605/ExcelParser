package com.rakbank.excelparser.model;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

@Data
@Slf4j
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
