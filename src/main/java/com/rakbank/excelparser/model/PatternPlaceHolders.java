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
    String pattern;
    String event_rq_template;
    String event_id;
}
