package org.jxls.issue.jxls30;

import org.jxls.area.XlsArea;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Created by Leonid Vysochyn on 09-Dec-15.
 */
public class Issue30Demo {
    public static void main(String[] args) throws ParseException, IOException {
        List<Employee> employees = generateSampleEmployeeData(1);
        List<Employee> employees2 = generateSampleEmployeeData(2);
        List<Employee> employees3 = generateSampleEmployeeData(2);
        List<Employee> employees4 = generateSampleEmployeeData(1);
        try(InputStream is = Issue30Demo.class.getResourceAsStream("issue30_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/issue30_output.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                XlsArea xlsArea = new XlsArea("Template!A1:I10", transformer);
                XlsArea employeeArea = new XlsArea("Template!A4:D4", transformer);
                EachCommand employeeEachCommand = new EachCommand("employee", "employees", employeeArea);

                XlsArea employeeArea2 = new XlsArea("Template!F4:I4", transformer);
                EachCommand employeeEachCommand2 = new EachCommand("employee2", "employees2", employeeArea2);

                XlsArea employeeArea3 = new XlsArea("Template!A8:D8", transformer);
                EachCommand employeeEachCommand3 = new EachCommand("employee3", "employees3", employeeArea3);

                XlsArea employeeArea4 = new XlsArea("Template!F8:I8", transformer);
                EachCommand employeeEachCommand4 = new EachCommand("employee4", "employees4", employeeArea4);

                xlsArea.addCommand("A4:D4", employeeEachCommand);
                xlsArea.addCommand("F4:I4", employeeEachCommand2);
                xlsArea.addCommand("A8:D8", employeeEachCommand3);
                xlsArea.addCommand("F8:I8", employeeEachCommand4);

                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("employees2", employees2);
                context.putVar("employees3", employees3);
                context.putVar("employees4", employees4);
                xlsArea.setFormulaProcessor(new StandardFormulaProcessor());
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                xlsArea.processFormulas();
                transformer.write();
            }
        }
    }

    private static List<Employee> generateSampleEmployeeData(int i) throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        if(i == 1) {
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
            employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
            employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25) );
            employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00) );
            employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15) );
            employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20) );
        }
        return employees;
    }
}
