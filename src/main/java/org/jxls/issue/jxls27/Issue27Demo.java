package org.jxls.issue.jxls27;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
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
 * Created by Leonid Vysochyn on 09-Dec-15.
 */
public class Issue27Demo {
    static Logger logger = LoggerFactory.getLogger(Issue27Demo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running ComplexFormulas issue#11 demo");
        List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
        try(InputStream is = Issue27Demo.class.getResourceAsStream("issue27_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue27_output.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                transformer.deleteSheet("AnyList");

                Context context = transformer.createInitialContext();
                context.putVar("employees", employees);
                /* some data addition to context */
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                for (Area xlsArea : xlsAreaList) {
                    // Exception here because deleted sheet is still in sheetMap but not in workbook
                    xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
                    xlsArea.setFormulaProcessor(new StandardFormulaProcessor()); /* FastFormulaProcessor */
                    xlsArea.processFormulas();
                }
                transformer.write();
            }
        }
    }
}
