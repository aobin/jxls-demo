package com.jxls.plus.demo;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12 5:05 PM
 */
public class MainDemo {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        EachIfCommandDemo.execute();
        EachIfXmlBuilderDemo.execute();
        FormulaExportDemo.execute();
        AreaListenerDemo.execute();
        MultipleSheetDemo.execute();
        UserCommandDemo.execute();
        XlsCommentBuilderDemo.execute();
        ImageDemo.execute();
    }
}
