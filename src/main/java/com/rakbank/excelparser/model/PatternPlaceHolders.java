package com.rakbank.excelparser.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

@Data
@Slf4j
@AllArgsConstructor
@NoArgsConstructor
public class PatternPlaceHolders {
    int sno;
    String pattern;
    String eventRqTemplate;
    String eventId;
}
