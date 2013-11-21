package com.jxls.plus.demo;

import com.jxls.plus.demo.guide.*;
import com.jxls.plus.transform.Transformer;
import com.jxls.plus.util.TransformerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12 5:05 PM
 */
public class MainDemo {
    public static void main(String[] args) throws Exception {
        ObjectCollectionDemo.main(args);
        ObjectCollectionJavaAPIDemo.main(args);
        ObjectCollectionFormulasDemo.main(args);
        ParameterizedFormulasDemo.main(args);
        ObjectCollectionXMLBuilderDemo.main(args);

        EachIfCommandDemo.execute();
        EachIfXmlBuilderDemo.execute();
        FormulaExportDemo.execute();

        MultipleSheetDemo.execute();
        XlsCommentBuilderDemo.execute();
        ImageDemo.execute();



        String transformerName = TransformerFactory.getTransformerName();

        if( TransformerFactory.POI_TRANSFORMER.equals( transformerName ) ){
            UserCommandExcelMarkupDemo.main(args);
            UserCommandDemo.execute();
            AreaListenerDemo.execute();
            StressXlsxDemo.executeStress1();
            StressXlsxDemo.executeStress2();
            SxssfDemo.executeStress1();
            SxssfDemo.executeStress2();
        }

        if( TransformerFactory.JEXCEL_TRANSFORMER.equals( transformerName)){
            JexcelAreaListenerDemo.execute();
            JexcelUserCommandExcelMarkupDemo.execute();
        }

        StressDemo.executeStress1();
        StressDemo.executeStress2();
    }
}
