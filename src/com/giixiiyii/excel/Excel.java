package com.giixiiyii.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public class Excel extends ListSupport<Excel.XSheet> {

    public XSheet newSheet() {
        return add(new XSheet("sheet-" + size())).get(size() - 1);
    }

    public byte[] getBytes() throws IOException {
        Workbook wb = new HSSFWorkbook();
        for (XSheet xSheet : getData()) {
            Sheet sheet = wb.createSheet(xSheet.getName());
            sheet.setDefaultColumnWidth(xSheet.getSchema().getWidth());

            int i = 0;
            createRow(wb, sheet, i++, xSheet.getSchema());
            for (XSheet.Record record : xSheet.getData())
                createRow(wb, sheet, i++, record);
        }
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        wb.write(stream);
        return stream.toByteArray();
    }

    void createRow(Workbook wb, Sheet sheet, int index, XSheet.Record record) {
        Row row = sheet.createRow(index);
        row.setHeightInPoints(record.getHeight());

        Font font = wb.createFont();
        font.setFontName(record.getFont());
        if (record instanceof XSheet.Schema)
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);

        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(record.getBackground());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        short border = HSSFCellStyle.BORDER_THIN;
        style.setBorderBottom(border);
        style.setBorderLeft(border);
        style.setBorderTop(border);
        style.setBorderRight(border);

        Cell cell;
        for (int i = 0; i < record.getData().size(); i++) {
            cell = row.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(record.getData().get(i));
        }
    }

    public class XSheet extends ListSupport<XSheet.Record> {
        public XSheet(String name) {
            name(name).schema(newSchema());
        }

        String name;
        Schema schema;

        public String getName() {
            return name;
        }

        public XSheet name(String name) {
            this.name = name;
            return this;
        }

        public Schema getSchema() {
            return schema;
        }

        public XSheet schema(Schema schema) {
            this.schema = schema;
            return this;
        }

        public Schema newSchema() {
            Schema scm = new Schema();
            schema(scm);
            return scm;
        }

        public Record newRecord() {
            return (Record) add(new Record()).get(size() - 1);
        }

        public class Schema extends Record {
            public Schema(int width) {
                super();
                width(width).background(IndexedColors.GREY_25_PERCENT.getIndex()).font("微软雅黑");
            }

            public Schema() {
                this(16);
            }

            int width;

            public int getWidth() {
                return width;
            }

            public Schema width(int width) {
                this.width = width;
                return this;
            }
        }

        //所有cell风格统一
        public class Record extends ListSupport<String> {
            public Record(int h, short b, String f, List data) {
                height((short) h);
                background(b);
                font(f);
                addAll(data);
            }

            public Record(List<String> data) {
                this(16, IndexedColors.WHITE.index, "宋体", data);
            }

            public Record(int height) {
                this(height, IndexedColors.WHITE.index, "宋体", null);
            }

            public Record() {
                this(16, IndexedColors.WHITE.index, "宋体", null);
            }

            short height;
            short background;
            String font;

            public short getHeight() {
                return height;
            }

            public Record height(short height) {
                this.height = height;
                return this;
            }

            public short getBackground() {
                return background;
            }

            public Record background(short background) {
                this.background = background;
                return this;
            }

            public String getFont() {
                return font;
            }

            public Record font(String font) {
                this.font = font;
                return this;
            }

        }
    }
}
