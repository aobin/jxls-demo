package org.jxls.demo;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.GridCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

/**
 * Created by Leonid Vysochyn on 30-Jun-15.
 */
public class GridCommandDemo {
    static Logger logger = LoggerFactory.getLogger(GridCommandDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Grid command demo");
        List<Employee> employees = generateSampleEmployeeData();
        InputStream is = GridCommandDemo.class.getResourceAsStream("grid_template.xls");
        OutputStream os = new FileOutputStream("target/grid_output.xls");
        Transformer transformer = TransformerFactory.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new Context();
        context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment"));
        List<List<Object>> data = new ArrayList<>();
        for(Employee employee : employees){
            data.add( convertEmployeeToList(employee));
        }
        context.putVar("data", data);
        xlsArea.applyAt(new CellRef("Sheet1!A1"), context);

        GridCommand gridCommand = (GridCommand) xlsArea.getCommandDataList().get(0).getCommand();
        gridCommand.setProps("name,payment,birthDate");
        context.putVar("headers", Arrays.asList("Name", "Payment", "Birthday"));
        context.putVar("data", employees);
        xlsArea.applyAt(new CellRef("Sheet2!A1"), context);

        transformer.write();
        is.close();
        os.close();
    }

    private static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add(new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25));
        employees.add(new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00));
        employees.add(new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15));
        employees.add(new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20));
        return employees;
    }

    private static List<Object> convertEmployeeToList(Employee employee){
        List<Object> list = new ArrayList<>();
        list.add(employee.getName());
        list.add(employee.getBirthDate());
        list.add(employee.getPayment());
        return list;
    }
}
