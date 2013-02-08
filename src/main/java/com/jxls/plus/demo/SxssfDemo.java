package com.jxls.plus.demo;

import com.jxls.plus.area.Area;
import com.jxls.plus.builder.AreaBuilder;
import com.jxls.plus.builder.xls.XlsCommentAreaBuilder;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.demo.model.Employee;
import com.jxls.plus.transform.Transformer;
import com.jxls.plus.transform.poi.PoiContext;
import com.jxls.plus.transform.poi.PoiTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
    static Logger logger = LoggerFactory.getLogger(SxssfDemo.class);

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing Stress demo");
        executeStress1();
//        executeStress2();
    }

    public static void executeStress1() throws IOException, InvalidFormatException {
        System.out.println(System.getProperty("java.io.tmpdir"));
        logger.info("Generating employees..");
        List<Employee> employees = Employee.generate(110);
        logger.info("Created " + employees.size() + " employees");
        InputStream is = SxssfDemo.class.getResourceAsStream("stress1.xlsx");
        assert is != null;
        Workbook workbook = WorkbookFactory.create(is);
        Transformer transformer = PoiTransformer.createTransformer(workbook, true);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new PoiContext();
        context.putVar("employees", employees);
        long startTime = System.nanoTime();
        xlsArea.applyAt(new CellRef("NewSheet!A1"), context);
        xlsArea.processFormulas();
        long endTime = System.nanoTime();
        System.out.println("Stress1 time (s): " + (endTime - startTime)/1000000000);
//        workbook.removeSheetAt(0);
        OutputStream os = new FileOutputStream("target/sxxfstress1_output.xlsx");
        ((PoiTransformer)transformer).getWorkbook().write(os);
        is.close();
        os.close();
    }

//    public static void executeStress2() throws IOException, InvalidFormatException {
//        logger.info("Generating departments..");
//        List<Department> departments = Department.generate(100, 500);
//        logger.info("Created " + departments.size() + " departments");
//        InputStream is = SxssfDemo.class.getResourceAsStream("stress2.xls");
//        assert is != null;
//        Workbook workbook = WorkbookFactory.create(is);
//        Transformer transformer = PoiTransformer.createTransformer(workbook);
//        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
//        List<Area> xlsAreaList = areaBuilder.build();
//        Area xlsArea = xlsAreaList.get(0);
//        Context context = new PoiContext();
//        context.putVar("departments", departments);
//        long startTime = System.nanoTime();
//        xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
//        xlsArea.processFormulas();
//        long endTime = System.nanoTime();
//        System.out.println("Stress1 time (s): " + (endTime - startTime)/1000000000);
//        workbook.removeSheetAt(0);
//        OutputStream os = new FileOutputStream("target/stress2_output.xls");
//        workbook.write(os);
//        is.close();
//        os.close();
//    }

}
