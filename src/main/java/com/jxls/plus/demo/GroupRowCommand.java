package com.jxls.plus.demo;

import com.jxls.plus.area.Area;
import com.jxls.plus.command.AbstractCommand;
import com.jxls.plus.command.Command;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.common.Size;
import com.jxls.plus.transform.poi.PoiTransformer;
import com.jxls.plus.util.Util;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * An implementation of a Command for row grouping
 * @author Leonid Vysochyn
 *         Date: 2/21/12
 */
public class GroupRowCommand extends AbstractCommand {
    Area area;
    String collapseIf;

    public String getName() {
        return "groupRow";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        Size resultSize = area.applyAt(cellRef, context);
        if( resultSize.equals(Size.ZERO_SIZE)) return resultSize;
        PoiTransformer transformer = (PoiTransformer) area.getTransformer();
        Workbook workbook = transformer.getWorkbook();
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        int startRow = cellRef.getRow();
        int endRow = cellRef.getRow() + resultSize.getHeight() - 1;
        sheet.groupRow(startRow, endRow);
        if( collapseIf != null && collapseIf.trim().length() > 0){
            boolean collapseFlag = Util.isConditionTrue(collapseIf, context);
            sheet.setRowGroupCollapsed(startRow, collapseFlag);
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
