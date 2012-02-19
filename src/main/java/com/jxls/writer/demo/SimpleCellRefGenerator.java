package com.jxls.writer.demo;

import com.jxls.writer.command.CellRefGenerator;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;

/**
 * @author Leonid Vysochyn
 */
public class SimpleCellRefGenerator implements CellRefGenerator{
    public CellRef generateCellRef(int index, Context context) {
        return new CellRef("sheet" + index + "!B2");
    }
}
