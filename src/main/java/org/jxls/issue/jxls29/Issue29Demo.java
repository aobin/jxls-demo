package org.jxls.issue.jxls29;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Leonid Vysochyn on 07-Dec-15.
 */
public class Issue29Demo {
    public static void main(String[] args) throws IOException {
        List<String> columns = new ArrayList<>();
        columns.add("a");
        columns.add("b");
        columns.add("c");
        columns.add("d");
        columns.add("e");
        columns.add("f");
        columns.add("g");
        Context data = new Context();
        data.putVar("columns", columns);
        try(InputStream is = Issue29Demo.class.getResourceAsStream("issue29_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue29_output.xls")) {
                JxlsHelper.getInstance().processTemplate(is, os, data);
            }
        }
    }
}
