package com.jxls.plus.demo;

import com.jxls.plus.command.CellRefGenerator;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;

/**
 * @author Leonid Vysochyn
 */
public class SimpleCellRefGenerator implements CellRefGenerator {
    public CellRef generateCellRef(int index, Context context) {
        return new CellRef("sheet" + index + "!B2");
    }
}
