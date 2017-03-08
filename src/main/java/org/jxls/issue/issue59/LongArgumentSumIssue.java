package org.jxls.issue.issue59;

import org.jxls.common.Context;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

public class LongArgumentSumIssue {
    private static final Logger logger = LoggerFactory.getLogger(LongArgumentSumIssue.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Long argument sum demo");
        List<Integer> values = generateLargeListOfNumbers(1000);
        try (InputStream is = LongArgumentSumIssue.class.getResourceAsStream("long-sum-argument-template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/long-sum-argument-output.xlsx")) {
                Context context = new Context();
                context.putVar("values", values);
                JxlsHelper.getInstance().setFormulaProcessor(new StandardFormulaProcessor()).
                        processTemplate(is, os, context);
            }
        }
    }

    private static List<Integer> generateLargeListOfNumbers(int count) {
        List<Integer> result = new ArrayList<>();
        for (int i = 1; i <= count; i++) {
            result.add(i);
        }
        return result;
    }
}
