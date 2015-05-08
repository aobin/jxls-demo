package org.jxls.demo;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.model.Department;
import org.jxls.demo.model.Employee;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class SxssfDemo {
    public static final int EMPLOYEE_COUNT = 30000;
    public static final int DEPARTMENT_COUNT = 100;
    public static final int DEP_EMPLOYEE_COUNT = 500;
    static Logger logger = LoggerFactory.getLogger(SxssfDemo.class);

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Entering Stress Sxssf demo");
        executeStress1();
        executeStress2();
    }

    public static void executeStress1() throws IOException, InvalidFormatException {
        logger.info("Running Stress Sxssf demo 1");
        logger.info("Generating " + EMPLOYEE_COUNT + " employees..");
        List<Employee> employees = Employee.generate(EMPLOYEE_COUNT);
        logger.info("Created " + employees.size() + " employees");
        InputStream is = SxssfDemo.class.getResourceAsStream("stress1.xlsx");
        assert is != null;
        Workbook workbook = WorkbookFactory.create(is);
        Transformer transformer = PoiTransformer.createSxssfTransformer(workbook);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new PoiContext();
        context.putVar("employees", employees);
        long startTime = System.nanoTime();
        xlsArea.applyAt(new CellRef("NewSheet!A1"), context);
//        xlsArea.processFormulas();
        long endTime = System.nanoTime();
        System.out.println("Stress Sxssf demo 1 time (s): " + (endTime - startTime)/1000000000);
        OutputStream os = new FileOutputStream("target/sxssf_stress1_output.xlsx");
        ((PoiTransformer)transformer).getWorkbook().write(os);
        is.close();
        os.close();
    }

    public static void executeStress2() throws IOException, InvalidFormatException {
        logger.info("Running Stress Sxssf demo 2");
        logger.info("Generating " + DEPARTMENT_COUNT + " departments with " + DEP_EMPLOYEE_COUNT + " employees in each");
        List<Department> departments = Department.generate(DEPARTMENT_COUNT, DEP_EMPLOYEE_COUNT);
        logger.info("Created " + departments.size() + " departments");
        InputStream is = SxssfDemo.class.getResourceAsStream("stress2.xlsx");
        assert is != null;
        Workbook workbook = WorkbookFactory.create(is);
        // setting rowAccessWindowSize to 600 to be able to process static cells in a single iteration
        Transformer transformer = PoiTransformer.createSxssfTransformer(workbook, 600, true);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new PoiContext();
        context.putVar("departments", departments);
        long startTime = System.nanoTime();
        xlsArea.applyAt(new CellRef("NewSheet!A1"), context);
//        xlsArea.processFormulas();
        long endTime = System.nanoTime();
        System.out.println("Stress Sxssf demo 2 time (s): " + (endTime - startTime)/1000000000);
        OutputStream os = new FileOutputStream("target/sxssf_stress2_output.xlsx");
        ((PoiTransformer)transformer).getWorkbook().write(os);
        is.close();
        os.close();
    }

}
