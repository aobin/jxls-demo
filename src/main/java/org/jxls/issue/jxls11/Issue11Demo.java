package org.jxls.issue.jxls11;

import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

/**
 * Created by Leonid Vysochyn on 07-Aug-15.
 */
public class Issue11Demo {
    static Logger logger = LoggerFactory.getLogger(Issue11Demo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running ComplexFormulas issue#11 demo");
        List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
        try(InputStream is = Issue11Demo.class.getResourceAsStream("issue11_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue11_output.xls")) {
                Context context = new Context();
                context.putVar("activities", employees);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }
}
