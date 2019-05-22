# excel
## 导出excel

```
Excel e = new Excel();
XSheet sheet;

sheet = e.newSheet().name("基本信息");
sheet.getSchema().add("学号").add("姓名");
sheet.newRecord().add("2010000000").add("Foo");
sheet.newRecord().add("2010000001").add("Bar");

sheet = e.newSheet();
sheet.getSchema().add("foo").add("bar");
sheet.newRecord().add("0").add("1");
sheet.newRecord().add("2").add("4");

export(e.getBytes());
```
## 导入excel
```
byte[] bytes = import2Bytes();
Excel excel = Excel.fromBytes(bytes);

//        sheet   row    cell
log(excel.get(0).get(1).get(1));

excel.forEach(sheet -> {
    log("----------------------");
    log(sheet.getName());
    log(sheet.getSchema());
    sheet.forEach(this::log);
});

/**
* 输出：
* Bar
* ----------------------
* 基本信息
* [学号, 姓名]
* [2010000000, Foo]
* [2010000001, Bar]
* ----------------------
* sheet-1
* [foo, bar]
* [0, 1]
* [2, 4]
```