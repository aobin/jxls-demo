package com.jxls.plus.demo;

import com.jxls.plus.area.Area;
import com.jxls.plus.builder.AreaBuilder;
import com.jxls.plus.builder.xls.XlsCommentAreaBuilder;
import com.jxls.plus.builder.xml.XmlAreaBuilder;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.demo.model.Employee;
import com.jxls.plus.transform.Transformer;
import com.jxls.plus.transform.jexcel.JexcelTransformer;
import com.jxls.plus.transform.poi.PoiContext;
import com.jxls.plus.transform.poi.PoiTransformer;
import com.jxls.plus.util.TransformerFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
 * @author Leonid Vysochyn
 */
public class UserCommandExcelMarkupDemo {
    static Logger logger = LoggerFactory.getLogger(UserCommandExcelMarkupDemo.class);
    private static String template = "user_command_markup_template.xls";
    private static String output = "target/user_command_markup_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException, ParseException {
        logger.info("Executing User Command Markup demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException, ParseException {
        logger.info("Running UserCommandExcelMarkupDemo");
        List<Employee> employees = generateSampleEmployeeData();
        logger.info("Opening input stream");
        InputStream is = UserCommandExcelMarkupDemo.class.getResourceAsStream(template);
        OutputStream os = new FileOutputStream(output);
        Transformer transformer = TransformerFactory.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        if( transformer instanceof PoiTransformer){
            XlsCommentAreaBuilder.addCommandMapping("groupRow", GroupRowCommand.class);
        }else{
            XlsCommentAreaBuilder.addCommandMapping("groupRow", JexcelGroupRowCommand.class);
        }
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new Context();
        context.putVar("employees", employees);
        xlsArea.applyAt(new CellRef("Result!A1"), context);
        transformer.write();
        is.close();
        logger.info("Finished UserCommandExcelMarkupDemo");
    }

    private static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500d, 0.15) );
        employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300d, 0.25) );
        employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500d, 0.00) );
        employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700d, 0.15) );
        employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800d, 0.20) );
        return employees;
    }
}
