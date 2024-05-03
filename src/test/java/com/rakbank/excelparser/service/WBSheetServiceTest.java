package com.rakbank.excelparser.service;

import com.rakbank.excelparser.model.Content;
import com.rakbank.excelparser.model.PatternPlaceHolders;
import com.rakbank.excelparser.serviceImpl.WBSheetServiceImpl;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

public class WBSheetServiceTest {

    @Test
    public void testExtractValuesWithDollar() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello $name_id$");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;", result.getEventRqTemplate());
    }
    @Test
    public void testExtractValuesWithPercentile() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello %name_id%");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesWithHash() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello #name_id#");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesWithFlowerBraces() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello {name_id}");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesWithFlowerBracesUnderscore() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello {Branch_DROPDOWN_BRANCH} abcd");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcd", result.getPattern());
        Assertions.assertEquals("Branch_DROPDOWN_BRANCH:param0;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesWithSquareBracesUnderscore() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello [name_id]");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;", result.getEventRqTemplate());
    }

    @Test
    public void textExtractValuesType1() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Your request $$Prospect_ID# for a credit card is successfully submitted . We will contact you in 2 business days to assist you with this request. Call us on 04 2130000 for any queries.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Your request (.*) for a credit card is successfully submitted . We will contact you in 2 business days to assist you with this request. Call us on 04 2130000 for any queries.", result.getPattern());
        Assertions.assertEquals("Prospect_ID:param0;", result.getEventRqTemplate());
    }

    @Test
    public void textExtractValuesType2() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("We regret that your request #Prospect_ID# for a #Product_Name# is not approved due to the Bank's policy. Call us on 04 2130000 for any queries");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("We regret that your request (.*) for a (.*) is not approved due to the Bank's policy. Call us on 04 2130000 for any queries", result.getPattern());
        Assertions.assertEquals("Prospect_ID:param0;Product_Name:param1;", result.getEventRqTemplate());
    }

    @Test
    public void textExtractValuesType3() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Collect your Internet banking ID & Password from {Branch_DROPDOWN_BRANCH} Branch within {Day_DROPDOWN_DAY} day(s) between 8AM-3PM Mon to Thur & Sat, 7:30AM-12:30PM Fri");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Collect your Internet banking ID & Password from (.*) Branch within (.*) day(s) between 8AM-3PM Mon to Thur & Sat, 7:30AM-12:30PM Fri", result.getPattern());
        Assertions.assertEquals("Branch_DROPDOWN_BRANCH:param0;Day_DROPDOWN_DAY:param1;", result.getEventRqTemplate());
    }



    @Test
    public void testExtractValuesWithoutSpecialCharacters() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("No special characters");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("No special characters", result.getPattern());
        Assertions.assertEquals("", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesTildeAndBackslash() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello `name_id` abcdef ~otp~");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesAtTheRateOf() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello @name_id@ abcdef @@otp@@@");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesPowerSymbol() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello ^name_id^ abcdef ^^otp^^^");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesRoundBraces() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello (name_id) abcdef (((otp))");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingPlus() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello +name_id+ abcdef ++otp++");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingHyphen() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello -name_id- abcdef --otp---");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingAmpersand() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello &name_id& abcdef &&otp&&&");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name_id:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingS() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello day(s) abcdef [otp]]");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello day(s) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingEquals() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello ==name=== abcdef [otp]]");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingLessThanEqualTo() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello <<name>>> abcdef <otp>");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*)", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingGreaterThanEqualTo() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello >>name<<< abcdef >otp<> gef.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*) gef.", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingQuestionMark() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello ??name?? abcdef ?otp?? gef.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*) gef.", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingBackSlash() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello /name// abcdef /otp// gef.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*) gef.", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingForwardSlash() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello \\name\\ abcdef \\otp\\ gef.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*) gef.", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingMultiply() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello **name** abcdef *otp*** gef.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*) gef.", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingPipe() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello ||name| abcdef |otp||| gef.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*) gef.", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }

    @Test
    public void testExtractValuesUsingColon() {
        WBSheetServiceImpl service = new WBSheetServiceImpl();
        Content content = new Content();
        content.setSmsTemplate("Hello ::name: abcdef :otp::: gef.");

        PatternPlaceHolders result = service.extractValues(content);

        Assertions.assertEquals("Hello (.*) abcdef (.*) gef.", result.getPattern());
        Assertions.assertEquals("name:param0;otp:param1;", result.getEventRqTemplate());
    }


    // Add more test cases as needed to cover different scenarios
}