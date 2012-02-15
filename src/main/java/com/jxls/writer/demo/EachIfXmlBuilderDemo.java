package com.jxls.writer.demo;

import com.jxls.writer.area.Area;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.builder.AreaBuilder;
import com.jxls.writer.builder.xml.XlsAreaXmlBuilder;
import com.jxls.writer.common.Context;
import com.jxls.writer.demo.model.Department;
import com.jxls.writer.transform.Transformer;
import com.jxls.writer.transform.poi.PoiTransformer;
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
 *         Date: 2/14/12 3:59 PM
 */
public class EachIfXmlBuilderDemo {
    static Logger logger = LoggerFactory.getLogger(EachIfCommandDemo.class);
    private static String template = "each_if_demo.xls";
    private static String xmlConfig = "each_if_demo.xml";
    private static String output = "target/each_if_xml_builder_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing Each,If XML builder demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        InputStream is = EachIfCommandDemo.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer poiTransformer = PoiTransformer.createTransformer(workbook);
        System.out.println("Creating areas");
        InputStream configInputStream = EachIfXmlBuilderDemo.class.getResourceAsStream(xmlConfig);
        AreaBuilder areaBuilder = new XlsAreaXmlBuilder(poiTransformer);
        List<Area> xlsAreaList = areaBuilder.build(configInputStream);
        Area xlsArea = xlsAreaList.get(0);
        Area xlsArea2 = xlsAreaList.get(1);
        Context context = new Context();
        context.putVar("departments", departments);
        logger.info("Applying first area at cell " + new CellRef("Down!A1"));
        xlsArea.applyAt(new CellRef("Down!A1"), context);
        xlsArea.processFormulas();
        logger.info("Applying second area at cell " + new CellRef("Right!A1"));
        poiTransformer.resetTargetCellRefs();
        xlsArea2.applyAt(new CellRef("Right!A1"), context);
        xlsArea2.processFormulas();
        logger.info("Complete");
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        logger.info("written to file");
        is.close();
        os.close();
    }

}
