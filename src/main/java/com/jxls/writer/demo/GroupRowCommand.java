package com.jxls.writer.demo;

import com.jxls.writer.area.Area;
import com.jxls.writer.command.AbstractCommand;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.Size;
import com.jxls.writer.transform.poi.PoiTransformer;
import com.jxls.writer.util.Util;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Leonid Vysochyn
 *         Date: 2/21/12 4:07 PM
 */
public class GroupRowCommand extends AbstractCommand{
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
    public void addArea(Area area) {
        super.addArea(area);
        this.area = area;
    }

    public void setCollapseIf(String collapseIf) {
        this.collapseIf = collapseIf;
    }
}
