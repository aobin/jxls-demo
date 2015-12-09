package org.jxls.issue.jxls28;

import org.jxls.area.XlsArea;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.SimpleCellRefGenerator;
import org.jxls.demo.model.Department;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * Created by Leonid Vysochyn on 07-Dec-15.
 */
public class Issue28Demo {
    public static void main(String[] args) throws IOException {
        excelExport();
    }

    public static void excelExport() throws IOException {
        List<Department> departments = Department.generate(10, 20, 15);
        try (InputStream is = Issue28Demo.class.getResourceAsStream("issue28_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue28_result.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                XlsArea xlsArea = new XlsArea("Template!A1:M17", transformer);
                XlsArea departmentArea = new XlsArea("Template!A2:M16", transformer);
                EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea, new SimpleCellRefGenerator());

                XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
                Command employeeEachCommand = new EachCommand("employee", "department.staff", employeeArea);
                departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);

                XlsArea employeeArea2 = new XlsArea("Template!H9:M9", transformer);
                Command employeeEachCommand2 = new EachCommand("employee2", "department.staff2", employeeArea2);
                departmentArea.addCommand(new AreaRef("Template!H9:M9"), employeeEachCommand2);

                XlsArea employeeArea3 = new XlsArea("Template!A14:F14", transformer);
                Command employeeEachCommand3 = new EachCommand("employee3", "department.staff", employeeArea3);
                departmentArea.addCommand(new AreaRef("Template!A14:F14"), employeeEachCommand3);

                XlsArea employeeArea4 = new XlsArea("Template!H14:M14", transformer);
                Command employeeEachCommand4 = new EachCommand("employee4", "department.staff2", employeeArea4);
                departmentArea.addCommand(new AreaRef("Template!H14:M14"), employeeEachCommand4);

                xlsArea.addCommand(new AreaRef("Template!A2:M16"), departmentEachCommand);
                Context context = new Context();
                context.putVar("departments", departments);
                xlsArea.applyAt(new CellRef("Template!A1"), context);
                xlsArea.processFormulas();
                transformer.write();
            } catch (Exception e) {
            }
        }
    }
}
