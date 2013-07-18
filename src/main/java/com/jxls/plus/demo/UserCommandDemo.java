package com.jxls.plus.demo;

import com.jxls.plus.area.Area;
import com.jxls.plus.builder.AreaBuilder;
import com.jxls.plus.builder.xml.XmlAreaBuilder;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.demo.model.Department;
import com.jxls.plus.transform.Transformer;
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
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/21/12 4:30 PM
 */
public class UserCommandDemo {
    static Logger logger = LoggerFactory.getLogger(UserCommandDemo.class);
    private static String template = "each_if_demo.xls";
    private static String xmlConfig = "user_command_demo.xml";
    private static String output = "target/user_command_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing User Command demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        InputStream is = EachIfCommandDemo.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer transformer = PoiTransformer.createTransformer(workbook);
        System.out.println("Creating areas");
        InputStream configInputStream = UserCommandDemo.class.getResourceAsStream(xmlConfig);
        AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new Context();
        context.putVar("departments", departments);
        logger.info("Applying area at cell " + new CellRef("Down!A1"));
        xlsArea.applyAt(new CellRef("Down!A1"), context);
        xlsArea.processFormulas();
        logger.info("Complete");
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        logger.info("Written to file");
        is.close();
        os.close();
    }

}
