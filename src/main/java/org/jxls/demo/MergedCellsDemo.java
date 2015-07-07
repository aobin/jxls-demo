package org.jxls.demo;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.model.Department;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * Created by Leonid Vysochyn on 6/30/2015.
 * todo: improve each command to be able to set merge cells
 */
public class MergedCellsDemo  {
    static Logger logger = LoggerFactory.getLogger(MergedCellsDemo.class);
    private static String template = "merged_cells_demo.xls";
    private static String output = "target/merged_cells_output.xls";

    public static void main(String[] args) throws IOException {
        logger.info("Running merged cells demo");
        execute();
    }

    public static void execute() throws IOException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        InputStream is = XlsCommentBuilderDemo.class.getResourceAsStream(template);
        OutputStream os = new FileOutputStream(output);
        Transformer transformer = TransformerFactory.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = transformer.createInitialContext();
        context.putVar("departments", departments);
        logger.info("Applying area " + xlsArea.getAreaRef() + " at cell " + new CellRef("Sheet1!A1"));
        xlsArea.applyAt(new CellRef("Sheet1!A1"), context);
        xlsArea.processFormulas();
        xlsArea.reset();
        logger.info("Complete");
        transformer.write();
        logger.info("written to file");
        is.close();
        os.close();
    }
}
