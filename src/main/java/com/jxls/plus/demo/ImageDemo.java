package com.jxls.plus.demo;

import com.jxls.plus.area.XlsArea;
import com.jxls.plus.command.ImageCommand;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.common.ImageType;
import com.jxls.plus.transform.Transformer;
import com.jxls.plus.transform.poi.PoiTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

/**
 * @author Leonid Vysochyn
 */
public class ImageDemo {
    static Logger logger = LoggerFactory.getLogger(ImageDemo.class);
    private static String template = "image_demo.xlsx";
    private static String output = "target/image_output.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing Image demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        logger.info("Opening input stream");
        InputStream is = ImageDemo.class.getResourceAsStream(template);
        OutputStream os = new FileOutputStream(output);
        Transformer transformer = PoiTransformer.createTransformer(is, os);
        XlsArea xlsArea = new XlsArea("Sheet1!A1:N30", transformer);
        Context context = new Context();
        InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.jpg");
        byte[] imageBytes = IOUtils.toByteArray(imageInputStream);
        context.putVar("image", imageBytes);
        XlsArea imgArea = new XlsArea("Sheet1!A5:D15", transformer);
        xlsArea.addCommand("Sheet1!A4:D15", new ImageCommand("image", ImageType.JPEG).addArea(imgArea));
        xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
        transformer.write();
        logger.info("written to file");
        is.close();
        os.close();
    }
}
