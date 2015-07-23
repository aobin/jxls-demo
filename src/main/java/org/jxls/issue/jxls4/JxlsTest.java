package org.jxls.issue.jxls4;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;

/**
 *
 * @author pernik
 */
public class JxlsTest {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            InputStream is = (JxlsTest.class.getResourceAsStream("issue4_template.xlsx"));
            OutputStream os = new FileOutputStream("target/issue4_output.xlsx");
            Transformer transformer = TransformerFactory.createTransformer(is, os);
            AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
            List<Area> xlsAreaList = areaBuilder.build();
            Context context = transformer.createInitialContext();

            List<Company> c = new ArrayList<>();
            c.add(new Company(1L, "A").addBoy("Karl").addBoy("Joseph").addGirl("Janet").addGirl("Lian").addGirl("Josephine"));
            c.add(new Company(2L, "B").addBoy("Karl 2").addBoy("Joseph 2").addBoy("John").addGirl("Janet 2").addGirl("Lian 2").addGirl("Josephine 2").addGirl("Jane"));
            c.add(new Company(3L, "B").addBoy("Karl 3").addBoy("Joseph 3").addBoy("John").addGirl("Janet 3").addGirl("Lian 3").addGirl("Josephine 3"));
            context.putVar("companies", c);
            for (Area xlsArea : xlsAreaList) {
                xlsArea.applyAt(/*
                         * new CellRef("Result!A1")
                         */new CellRef(xlsArea.getStartCellRef().getCellName()), context);
            }
            transformer.write();
            is.close();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
