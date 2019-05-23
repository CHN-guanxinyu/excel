package io.github.tobetwo.excel;

import javafx.util.Pair;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class Excel extends ListSupport<Excel.XSheet> {

    public XSheet newSheet() {
        return append(new XSheet("sheet-" + size())).get(size() - 1);
    }

    public byte[] getBytes() throws IOException {
        Workbook wb = new HSSFWorkbook();
        if (isEmpty()) newSheet();
        forEach(createSheet(wb));
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        wb.write(stream);
        return stream.toByteArray();
    }

    Consumer<XSheet> createSheet(Workbook wb){
        return xSheet -> {
            Sheet sheet = wb.createSheet(xSheet.getName());
            sheet.setDefaultColumnWidth(xSheet.getSchema().getWidth());

            createRow(wb, sheet).accept(new Pair(-1, xSheet.getSchema()));
            zipWithIndex(xSheet).forEach(createRow(wb, sheet));
        };
    }

    Consumer<Pair<Integer, XSheet.Record>> createRow(Workbook wb, Sheet sheet) {
        return pair -> {
            Row row = sheet.createRow(pair.getKey() + 1);
            XSheet.Record record = pair.getValue();
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

            zipWithIndex(record).forEach(createCell(row, style));
        };
    }

    Consumer<Pair<Integer, String>> createCell(Row row, CellStyle style) {
        return pair -> {
            Cell cell = row.createCell(pair.getKey());
            cell.setCellStyle(style);
            cell.setCellValue(pair.getValue());
        };
    }

    List zipWithIndex(List list) {
        return IntStream.range(0, list.size()).mapToObj(i -> new Pair(i, list.get(i))).collect(Collectors.toList());
    }

    public static Excel fromBytes(byte[] bytes) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new ByteArrayInputStream(bytes));

        Excel excel = new Excel();
        Stream.iterate(0, n -> n + 1)
                .limit(wb.getNumberOfSheets())
                .map(wb::getSheetAt)
                .forEach(wbSheet -> {
                    XSheet sheet = excel.newSheet().name(wbSheet.getSheetName());
                    sink(wbSheet.getRow(wbSheet.getTopRow()), sheet.getSchema());
                    for (int i = wbSheet.getTopRow() + 1; i <= wbSheet.getLastRowNum(); i++)
                        sink(wbSheet.getRow(i), sheet.newRecord());
                });
        return excel;
    }

    static void sink(Row row, List sink) {
        row.forEach(cell -> sink.add(getStringVal(cell)));
    }

    static String getStringVal(Cell cell) {
        if (cell == null)
            return "";
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case HSSFCell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            case HSSFCell.CELL_TYPE_NUMERIC:
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            case HSSFCell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                return "";
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
            return append(new Record()).get(size() - 1);
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

        public class Record extends ListSupport<String> {
            public Record(Collection<String> data) {
                appendAll(data);
            }


            public Record() {
                this(null);
            }

            short height = 16;
            short background = IndexedColors.WHITE.index;
            String font = "宋体";

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
