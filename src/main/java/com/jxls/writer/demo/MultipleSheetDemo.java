package com.jxls.writer.demo;

import com.jxls.writer.area.XlsArea;
import com.jxls.writer.command.Command;
import com.jxls.writer.command.EachCommand;
import com.jxls.writer.command.IfCommand;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.demo.model.Department;
import com.jxls.writer.demo.model.Employee;
import com.jxls.writer.transform.Transformer;
import com.jxls.writer.transform.poi.PoiTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class MultipleSheetDemo {
    static Logger logger = LoggerFactory.getLogger(MultipleSheetDemo.class);
    private static String template = "each_if_demo.xls";
    private static String output = "target/multiple_sheet_demo_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing Multiple Sheet demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        InputStream is = EachIfCommandDemo.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer poiTransformer = PoiTransformer.createTransformer(workbook);
        System.out.println("Creating area");
        XlsArea xlsArea = new XlsArea("Template!A1:G15", poiTransformer);
        XlsArea departmentArea = new XlsArea("Template!A2:G12", poiTransformer);
        EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea, new SimpleCellRefGenerator());
        XlsArea employeeArea = new XlsArea("Template!A9:F9", poiTransformer);
        XlsArea ifArea = new XlsArea("Template!A18:F18", poiTransformer);
        IfCommand ifCommand = new IfCommand("employee.payment <= 2000",
                ifArea,
                new XlsArea("Template!A9:F9", poiTransformer));
        employeeArea.addCommand("Template!A9:F9", ifCommand);
        Command employeeEachCommand = new EachCommand( "employee", "department.staff", employeeArea);
        departmentArea.addCommand("Template!A9:F9", employeeEachCommand);
        xlsArea.addCommand("Template!A2:F12", departmentEachCommand);
        Context context = new Context();
        context.putVar("departments", departments);
        logger.info("Applying at cell Sheet!A1");
        xlsArea.applyAt(new CellRef("Sheet!A1"), context);
        xlsArea.processFormulas();
        logger.info("Complete");
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        logger.info("written to file");
        is.close();
        os.close();
    }

}
