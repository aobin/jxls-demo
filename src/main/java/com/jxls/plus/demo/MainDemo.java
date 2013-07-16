package com.jxls.plus.demo;

import com.jxls.plus.demo.guide.ObjectCollectionDemo;
import com.jxls.plus.demo.guide.ObjectCollectionFormulasDemo;
import com.jxls.plus.demo.guide.ParameterizedFormulasDemo;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.text.ParseException;

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12 5:05 PM
 */
public class MainDemo {
    public static void main(String[] args) throws IOException, InvalidFormatException, WriteException, BiffException, ParseException {
        ObjectCollectionDemo.main(args);
        ObjectCollectionFormulasDemo.main(args);
        ParameterizedFormulasDemo.main(args);
//        EachIfCommandDemo.execute();
//        EachIfXmlBuilderDemo.execute();
//        FormulaExportDemo.execute();
//        AreaListenerDemo.execute();
//        MultipleSheetDemo.execute();
//        UserCommandDemo.execute();
//        XlsCommentBuilderDemo.execute();
//        JexcelXlsCommentBuilderDemo.execute();
//        ImageDemo.execute();
//        StressDemo.executeStress1();
//        StressDemo.executeStress2();
//        StressXlsxDemo.executeStress1();
//        StressXlsxDemo.executeStress2();
//        SxssfDemo.executeStress1();
//        SxssfDemo.executeStress2();
//        JexcelStressDemo.executeStress1();
//        JexcelStressDemo.executeStress2();
    }
}
