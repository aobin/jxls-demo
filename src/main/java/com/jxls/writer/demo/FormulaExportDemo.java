package com.jxls.writer.demo;

import com.jxls.writer.Pos;
import com.jxls.writer.Size;
import com.jxls.writer.command.*;
import com.jxls.writer.demo.model.Department;
import com.jxls.writer.demo.model.Employee;
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

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12 4:38 PM
 */
public class FormulaExportDemo {
    static Logger logger = LoggerFactory.getLogger(FormulaExportDemo.class);
    private static String template = "formulas_demo.xlsx";
    private static String output = "target/formulas_demo_output.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        logger.info("Opening input stream");
        InputStream is = FormulaExportDemo.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer poiTransformer = PoiTransformer.createTransformer(workbook);
        BaseArea area = new BaseArea(new Pos("Sheet1",0,0), new Size(4,4), poiTransformer);
        //BaseArea sheet2Area = new BaseArea(new Pos(0,0,0), new Size(1,2), poiTransformer);
        Context context = new Context();
        //sheet2Area.applyAt(new Pos(1, 5, 1), context);
        area.applyAt(new Pos("Sheet1", 10, 5), context);
        area.processFormulas();
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        logger.info("written to file");
        is.close();
        os.close();
    }

}
