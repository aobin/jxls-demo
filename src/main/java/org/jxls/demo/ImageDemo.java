package org.jxls.demo;

import org.jxls.area.XlsArea;
import org.jxls.command.ImageCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ImageType;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.jxls.util.Util;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author Leonid Vysochyn
 */
public class ImageDemo {
    static Logger logger = LoggerFactory.getLogger(ImageDemo.class);
    private static String template = "image_demo.xls";
    private static String output = "target/image_output.xls";

    public static void main(String[] args) throws IOException {
        logger.info("Running Image demo");
        execute();
    }

    public static void execute() throws IOException {
        logger.info("Opening input stream");
        InputStream is = ImageDemo.class.getResourceAsStream(template);
        OutputStream os = new FileOutputStream(output);
        Transformer transformer = TransformerFactory.createTransformer(is, os);
        XlsArea xlsArea = new XlsArea("Sheet1!A1:N30", transformer);
        Context context = new Context();
        InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
        byte[] imageBytes = Util.toByteArray(imageInputStream);
        context.putVar("image", imageBytes);
        XlsArea imgArea = new XlsArea("Sheet1!A5:D15", transformer);
        xlsArea.addCommand("Sheet1!A4:D15", new ImageCommand("image", ImageType.PNG).addArea(imgArea));
        xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
        transformer.write();
        logger.info("written to file");
        is.close();
        os.close();
    }
}
