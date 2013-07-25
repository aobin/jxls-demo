package com.jxls.plus.demo.guide;

import com.jxls.plus.area.Area;
import com.jxls.plus.builder.AreaBuilder;
import com.jxls.plus.builder.xls.XlsCommentAreaBuilder;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.transform.Transformer;
import com.jxls.plus.transform.poi.PoiContext;
import com.jxls.plus.transform.poi.PoiTransformer;
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
 * Object collection output demo
 * @author Leonid Vysochyn
 */
public class ObjectCollectionFormulasDemo {
    static Logger logger = LoggerFactory.getLogger(ObjectCollectionFormulasDemo.class);

    public static void main(String[] args) throws ParseException, IOException, InvalidFormatException {
        logger.info("Running ObjectCollectionDemo");
        List<Employee> employees = generateSampleEmployeeData();
        InputStream is = ObjectCollectionFormulasDemo.class.getResourceAsStream("formulas_template.xls");
        OutputStream os = new FileOutputStream("target/formulas_output.xls");
        Transformer transformer = PoiTransformer.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new PoiContext();
        context.putVar("employees", employees);
        xlsArea.applyAt(new CellRef("Result!A1"), context);
        xlsArea.processFormulas();
        transformer.write();
        is.close();
        os.close();
        logger.info("Finished ObjectCollectionFormulasDemo");
    }

    private static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25) );
        employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00) );
        employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15) );
        employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20) );
        return employees;
    }
}
