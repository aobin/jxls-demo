package com.jxls.writer.demo;

import com.jxls.writer.area.XlsArea;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.command.*;
import com.jxls.writer.common.Context;
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
public class EachIfCommandDemo {
    static Logger logger = LoggerFactory.getLogger(EachIfCommandDemo.class);
    private static String template = "each_if_demo.xls";
    private static String output = "target/each_if_demo_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing Each,If command demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        List<Department> departments = createDepartments();
        logger.info("Opening input stream");
        InputStream is = EachIfCommandDemo.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer poiTransformer = PoiTransformer.createTransformer(workbook);
        System.out.println("Creating area");
        XlsArea xlsArea = new XlsArea("Template!A1:G15", poiTransformer);
        XlsArea departmentArea = new XlsArea("Template!A2:G12", poiTransformer);
        EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
        XlsArea employeeArea = new XlsArea("Template!A9:F9", poiTransformer);
        IfCommand ifCommand = new IfCommand("employee.payment <= 2000",
                new XlsArea("Template!A18:F18", poiTransformer),
                new XlsArea("Template!A9:F9", poiTransformer));
        employeeArea.addCommand("Template!A9:F9", ifCommand);
        Command employeeEachCommand = new EachCommand( "employee", "department.staff", employeeArea);
        departmentArea.addCommand("Template!A9:F9", employeeEachCommand);
        xlsArea.addCommand("Template!A2:F12", departmentEachCommand);
        Context context = new Context();
        context.putVar("departments", departments);
        logger.info("Applying at cell " + new CellRef("Down!A1"));
        xlsArea.applyAt(new CellRef("Down!A1"), context);
        xlsArea.processFormulas();
        logger.info("Setting EachCommand direction to Right");
        departmentEachCommand.setDirection(EachCommand.Direction.RIGHT);
        logger.info("Applying at cell " + new CellRef("Right!A1"));
        poiTransformer.resetTargetCellRefs();
        xlsArea.applyAt(new CellRef("Right!A1"), context);
        xlsArea.processFormulas();
        logger.info("Complete");
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        logger.info("written to file");
        is.close();
        os.close();
    }

    public static List<Department> createDepartments() {
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
        return departments;
    }

}
