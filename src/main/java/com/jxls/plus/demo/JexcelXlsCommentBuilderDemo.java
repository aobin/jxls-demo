package com.jxls.plus.demo;

import com.jxls.plus.area.Area;
import com.jxls.plus.builder.AreaBuilder;
import com.jxls.plus.builder.xls.XlsCommentAreaBuilder;
import com.jxls.plus.command.EachCommand;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.demo.model.Department;
import com.jxls.plus.transform.jexcel.JexcelContext;
import com.jxls.plus.transform.jexcel.JexcelTransformer;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class JexcelXlsCommentBuilderDemo {
    static Logger logger = LoggerFactory.getLogger(JexcelXlsCommentBuilderDemo.class);
    private static String template = "comment_markup_demo.xls";
    private static String output = "target/jexcel_comment_builder_output.xls";

    public static void main(String[] args) throws IOException, BiffException, WriteException {
        logger.info("Executing Jexcel XLS Comment builder demo");
        execute();
    }


    public static void execute() throws IOException, BiffException, WriteException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        InputStream is = JexcelXlsCommentBuilderDemo.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating JexcelTransformer");
        OutputStream os = new BufferedOutputStream(new FileOutputStream(output));
        JexcelTransformer transformer = JexcelTransformer.createTransformer(is, os);
        System.out.println("Creating areas");
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = new JexcelContext();
        context.putVar("departments", departments);
//        InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.jpg");
//        byte[] imageBytes = IOUtils.toByteArray(imageInputStream);
//        context.putVar("image", imageBytes);
        logger.info("Applying area " + xlsArea.getAreaRef() + " at cell " + new CellRef("Down!A1"));
        xlsArea.applyAt(new CellRef("Down!A1"), context);
        xlsArea.processFormulas();
        xlsArea.reset();
        EachCommand eachCommand = (EachCommand) xlsArea.findCommandByName("each").get(0);
        eachCommand.setDirection(EachCommand.Direction.RIGHT);
        logger.info("Applying area " + xlsArea.getAreaRef() + " at cell " + new CellRef("Right!A1"));
        xlsArea.applyAt(new CellRef("Right!A1"), context);
        xlsArea.processFormulas();
        logger.info("Removing template sheet");
        transformer.getWritableWorkbook().removeSheet(0);
        is.close();
        transformer.getWritableWorkbook().write();
        transformer.getWritableWorkbook().close();
        logger.info("Complete");
        logger.info("written to file");
    }
}
