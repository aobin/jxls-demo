package com.jxls.plus.demo.jexcel;

import com.jxls.plus.area.Area;
import com.jxls.plus.builder.AreaBuilder;
import com.jxls.plus.builder.xls.XlsCommentAreaBuilder;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.demo.StressDemo;
import com.jxls.plus.demo.model.Department;
import com.jxls.plus.demo.model.Employee;
import com.jxls.plus.transform.jexcel.JexcelContext;
import com.jxls.plus.transform.jexcel.JexcelTransformer;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class JexcelStressDemo {
    static Logger logger = LoggerFactory.getLogger(JexcelStressDemo.class);

    public static void main(String[] args) throws BiffException, IOException, WriteException {
        logger.info("Executing Jexcel Stress demo");
        executeStress1();
        executeStress2();
    }

    public static void executeStress1() throws IOException, WriteException, BiffException {
        logger.info("Generating employees..");
        List<Employee> employees = Employee.generate(30000);
        logger.info("Created " + employees.size() + " employees");
        InputStream is = StressDemo.class.getResourceAsStream("stress1.xls");
        assert is != null;
        OutputStream os = new BufferedOutputStream(new FileOutputStream("target/jexcel_stress1_output.xls"));
        JexcelTransformer transformer = JexcelTransformer.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new JexcelContext();
        context.putVar("employees", employees);
        long startTime = System.nanoTime();
        xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
        xlsArea.processFormulas();
        long endTime = System.nanoTime();
        System.out.println("Stress1 time (s): " + (endTime - startTime) / 1000000000);
        transformer.getWritableWorkbook().removeSheet(0);
        is.close();
        transformer.getWritableWorkbook().write();
        transformer.getWritableWorkbook().close();
        os.close();
    }

    public static void executeStress2() throws IOException, BiffException, WriteException {
        logger.info("Generating departments..");
        List<Department> departments = Department.generate(100, 500);
        logger.info("Created " + departments.size() + " departments");
        InputStream is = StressDemo.class.getResourceAsStream("stress2.xls");
        assert is != null;
        OutputStream os = new BufferedOutputStream(new FileOutputStream("target/jexcel_stress2_output.xls"));
        JexcelTransformer transformer = JexcelTransformer.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new JexcelContext();
        context.putVar("departments", departments);
        long startTime = System.nanoTime();
        xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
        xlsArea.processFormulas();
        long endTime = System.nanoTime();
        System.out.println("Stress2 time (s): " + (endTime - startTime) / 1000000000);
        transformer.getWritableWorkbook().removeSheet(0);
        transformer.getWritableWorkbook().write();
        transformer.getWritableWorkbook().close();
        is.close();
        os.close();
    }
}
