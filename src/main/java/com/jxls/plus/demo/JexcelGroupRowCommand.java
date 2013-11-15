package com.jxls.plus.demo;

import com.jxls.plus.area.Area;
import com.jxls.plus.command.AbstractCommand;
import com.jxls.plus.command.Command;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.common.Size;
import com.jxls.plus.transform.jexcel.JexcelTransformer;
import com.jxls.plus.util.Util;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 11/15/13
 */
public class JexcelGroupRowCommand extends AbstractCommand {
    static Logger logger = LoggerFactory.getLogger(JexcelGroupRowCommand.class);
    Area area;
    String collapseIf;

    public String getName() {
        return "groupRow";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        Size resultSize = area.applyAt(cellRef, context);
        if( resultSize.equals(Size.ZERO_SIZE)) return resultSize;
        int startRow = cellRef.getRow();
        int endRow = cellRef.getRow() + resultSize.getHeight() - 1;
        try{
            JexcelTransformer transformer = (JexcelTransformer) area.getTransformer();
            WritableWorkbook workbook = transformer.getWritableWorkbook();
            WritableSheet sheet = workbook.getSheet(cellRef.getSheetName());
            boolean collapseFlag = false;
            if( collapseIf != null && collapseIf.trim().length() > 0){
                collapseFlag = Util.isConditionTrue(collapseIf, context);
            }
            sheet.setRowGroup(startRow, endRow, collapseFlag);
        }catch(Exception e){
            logger.error("Failed to apply JexcelGroupRowCommand at " + cellRef, e);
        }
        return resultSize;
    }

    @Override
    public Command addArea(Area area) {
        super.addArea(area);
        this.area = area;
        return this;
    }

    public void setCollapseIf(String collapseIf) {
        this.collapseIf = collapseIf;
    }
}
