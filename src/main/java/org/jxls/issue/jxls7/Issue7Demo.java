package org.jxls.issue.jxls7;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiTransformer;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;

public class Issue7Demo {

    public static MainModel generateData() {
        MainModel mainModel = new MainModel();
        mainModel.setMonths(Arrays.asList(new String[] {"May", "June"}));
        mainModel.setUsers(new ArrayList<User>());
        for(int i = 1; i<=3; i++) {
            User user = new User();
            user.setAccount(i+100);
            user.setName("Person " + i);
            user.setDatas(new ArrayList<Data>());
            for(int j = 1; j<3; j++) {
                Data data = new Data();
                data.setValue(j);
                user.getDatas().add(data);
            }
            mainModel.getUsers().add(user);
        }
        return mainModel;
    }

    public static void main(String[] args) throws Exception {
        XlsCommentAreaBuilder.addCommandMapping("each", EachRightCommand.class);
        try(InputStream fis = Issue7Demo.class.getResourceAsStream("issue7_template.xlsx")) {
            try(FileOutputStream fos = new FileOutputStream("target/issue7_output.xlsx")) {

                Workbook workbook = WorkbookFactory.create(fis);
                PoiTransformer transformer = PoiTransformer.createTransformer(workbook);
                Area area = (new XlsCommentAreaBuilder(transformer)).build().get(0);
                Context ctx = new PoiContext();
                ctx.putVar("re", generateData());
                area.applyAt(new CellRef(workbook.getSheetAt(0).getSheetName() + "!A1"), ctx);
                area.processFormulas();
                workbook.write(fos);
            }
        }
    }
}
