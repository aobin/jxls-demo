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

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 1/30/12 12:15 PM
 */
public class NestedEachLoopSectionExport {
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
        System.out.println("Opening input stream");
        InputStream is = NestedEachLoopSectionExport.class.getResourceAsStream(template);
        assert is != null;
        System.out.println("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer poiTransformer = new PoiTransformer(workbook);
        System.out.println("Creating area");
        BaseArea baseArea = new BaseArea(new Cell(0,0), new Size(7, 15), poiTransformer);
        Command eachCommand = new EachCommand(new Size(6, 11), "department", "departments", new BaseArea(new Cell(0, 1), new Size(6, 11), poiTransformer));
        baseArea.addCommand(new Pos(0, 1), eachCommand);
        Context context = new Context();
        context.putVar("departments", departments);
        System.out.println("Applying at cell (0,0,1)");
        baseArea.applyAt(new Cell(0, 0, 1), context);
        System.out.println("Complete");
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        System.out.println("written to file");
        is.close();
        os.close();
    }
}
