package com.jxls.writer.demo;

import com.jxls.writer.Cell;
import com.jxls.writer.Pos;
import com.jxls.writer.Size;
import com.jxls.writer.command.*;
import com.jxls.writer.transform.Transformer;
import com.jxls.writer.transform.poi.PoiTransformer;
import com.jxls.writer.demo.model.Department;
import com.jxls.writer.demo.model.Employee;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 1/30/12 12:15 PM
 */
public class NestedEachLoopSectionExport {
    static Logger logger = LoggerFactory.getLogger(NestedEachLoopSectionExport.class);
    private static String template = "departments.xls";
    private static String output = "target/departments_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        List<Department> departments = new ArrayList<Department>();
        Department department = new Department("IT");
        Employee chief = new Employee("Derek", 35, 3000, 0.30);
        department.setChief(chief);
        department.addEmployee(new Employee("Elsa", 28, 1500, 0.15));
        department.addEmployee(new Employee("Oleg", 32, 2300, 0.25));
        department.addEmployee(new Employee("Neil", 34, 2500, 0.00));
        department.addEmployee(new Employee("Maria", 34, 1700, 0.15));
        department.addEmployee(new Employee("John", 35, 2800, 0.20));
        departments.add(department);
        department = new Department("HR");
        chief = new Employee("Betsy", 37, 2200, 0.30);
        department.setChief(chief);
        department.addEmployee(new Employee("Olga", 26, 1400, 0.20));
        department.addEmployee(new Employee("Helen", 30, 2100, 0.10));
        department.addEmployee(new Employee("Keith", 24, 1800, 0.15));
        department.addEmployee(new Employee("Cat", 34, 1900, 0.15));
        departments.add(department);
        department = new Department("BA");
        chief = new Employee("Wendy", 35, 2900, 0.35);
        department.setChief(chief);
        department.addEmployee(new Employee("Denise", 30, 2400, 0.20));
        department.addEmployee(new Employee("LeAnn", 32, 2200, 0.15));
        department.addEmployee(new Employee("Natali", 28, 2600, 0.10));
        department.addEmployee(new Employee("Martha", 33, 2150, 0.25));
        departments.add(department);
        logger.info("Opening input stream");
        InputStream is = NestedEachLoopSectionExport.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer poiTransformer = new PoiTransformer(workbook);
        System.out.println("Creating area");
        BaseArea baseArea = new BaseArea(new Pos(0, 0), new Size(7, 15), poiTransformer);
        BaseArea departmentArea = new BaseArea(new Pos(1, 0), new Size(7, 11), poiTransformer);
        EachCommand eachCommand = new EachCommand(new Size(7, 11), "department", "departments", departmentArea);
        BaseArea employeeArea = new BaseArea(new Pos(8, 0), new Size(6, 1), poiTransformer);
        IfCommand ifCommand = new IfCommand("employee.payment <= 2000", new Size(6,1),
                new BaseArea(new Pos(17, 0), new Size(6,1), poiTransformer),
                new BaseArea(new Pos(8, 0), new Size(6,1), poiTransformer));
        employeeArea.addCommand(new Pos(0, 0), ifCommand);
        Command employeeEachCommand = new EachCommand(new Size(6,1), "employee", "department.staff", employeeArea);
        departmentArea.addCommand(new Pos(7, 0), employeeEachCommand);
        baseArea.addCommand(new Pos(1, 0), eachCommand);
        Context context = new Context();
        context.putVar("departments", departments);
        logger.info("Applying at cell (0,1,0)");
        baseArea.applyAt(new Pos(1, 0, 0), context);
        logger.info("Setting EachCommand direction to Right");
        eachCommand.setDirection(EachCommand.Direction.RIGHT);
        logger.info("Applying at cell (2,0,0)");
        baseArea.applyAt(new Pos(2, 0, 0), context);
        logger.info("Complete");
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        logger.info("written to file");
        is.close();
        os.close();
    }

}
