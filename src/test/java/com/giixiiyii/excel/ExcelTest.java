package com.giixiiyii.excel;

import com.giixiiyii.excel.Excel.XSheet;
import org.junit.Test;

import java.io.*;

public class ExcelTest {
    @Test
    public void export() throws IOException {
        Excel e = new Excel();
        XSheet sheet;

        sheet = e.newSheet().name("基本信息");
        sheet.getSchema().append("学号").append("姓名");
        sheet.newRecord().append("2015201304").append("关新宇");
        sheet.newRecord().append("2015201306").append("李涛");

        sheet = e.newSheet();
        sheet.getSchema().append("foo").append("bar");
        sheet.newRecord().append("0").append("1");
        sheet.newRecord().append("2").append("4");

        export(e.getBytes(), "D:/tmp/excel.xls");
    }

    void export(byte[] bytes, String filePath) throws IOException {
        OutputStream os = new FileOutputStream(new File(filePath));
        os.write(bytes);
        os.close();
    }

    @Test
    public void importTest() throws IOException {
        byte[] bytes = import2Bytes();
        Excel excel = Excel.fromBytes(bytes);

        //         sheet   row    cell
        log(excel.get(0).get(1).get(1));

        excel.forEach(sheet -> {
            log("----------------------");
            log(sheet.getName());
            log(sheet.getSchema());
            sheet.forEach(this::log);
        });
    }

    void log(Object o) {
        System.out.println(o);
    }

    byte[] import2Bytes() throws IOException {
        File file = new File("D:/tmp/excel.xls");
        byte[] bytes = new byte[(int) file.length()];
        InputStream is = new FileInputStream(file);
        is.read(bytes);
        return bytes;
    }
}
