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
Excel e = Excel.fromBytes(bytes);
XSheet sheet;
sheet = e.get(0);
List schema = sheet.getSchema().getData(); //List(学号，姓名)
sheet.get(0).get(0); //2010000000
sheet.get(0).get(1); //Foo
sheet.get(1).get(1); //Bar
```