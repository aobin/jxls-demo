package org.jxls.demo;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

/**
 * Created by Leonid Vysochyn on 22-Jul-15.
 */
public class CustomExpressionNotationDemo {
    static Logger logger = LoggerFactory.getLogger(CustomExpressionNotationDemo.class);

    private static final String TEMPLATE = "custom_expression_notation_template.xlsx";
    private static final String OUTPUT = "target/custom_expression_notation_output.xlsx";

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Custom Expression Notation demo");
        List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
        InputStream is = CustomExpressionNotationDemo.class.getResourceAsStream(TEMPLATE);
        OutputStream os = new FileOutputStream(OUTPUT);
        Transformer transformer = TransformerFactory.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        transformer.getTransformationConfig().buildExpressionNotation("[[", "]]");
        Context context = new Context();
        context.putVar("employees", employees);
        xlsArea.applyAt(new CellRef("Template!A1"), context);
        transformer.write();
        is.close();
        os.close();
    }

}
