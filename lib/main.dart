
import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_core/theme.dart';
import 'package:syncfusion_flutter_datagrid/datagrid.dart';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart' hide Column, Alignment;

void main() => runApp(MyApp());

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter ShowCase',
      theme: ThemeData(
        primaryColor: Color(0xffEE5366),
      ),
      debugShowCheckedModeBanner: false,
      home: Scaffold(
        body: MailPage(),
      ),
    );
  }
}

class MailPage extends StatefulWidget {
  @override
  _MailPageState createState() => _MailPageState();
}

class _MailPageState extends State<MailPage> {

  List<Employee> employees = <Employee>[];
  late EmployeeDataSource employeeDataSource;
  List<String> alphabetCharacters = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P"];

  @override
  void initState() {
    super.initState();
    employees = getEmployeeData();
    employeeDataSource = EmployeeDataSource(employeeData: employees);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Colors.white,
      body: Column(
        children: [
          ElevatedButton(
            onPressed: () {
              generateExcel();
            },
            style: ButtonStyle(
              backgroundColor: MaterialStateProperty.all<Color>(
                  const Color(0xff333399)),
              shape: MaterialStateProperty.all<
                  RoundedRectangleBorder>(
                RoundedRectangleBorder(
                  side: const BorderSide(
                    color: Color(0xff333399),
                  ),
                  borderRadius: BorderRadius.circular(10),
                ),
              ),
            ),
            child: const Text(
              "İndir",
              style: TextStyle(
                  color: Colors.white,
                  fontWeight: FontWeight.bold),
            ),
          ),
          SfDataGridTheme(
            data: SfDataGridThemeData(
              headerColor: Colors.red,
            ),
            child: SfDataGrid(
              onSelectionChanged: (addedRows, removedRows) {
                print(addedRows);
                print(removedRows);
              },
              source: employeeDataSource,
              columnWidthMode: ColumnWidthMode.fill,
              allowSorting: true,
              selectionMode: SelectionMode.multiple,
              showCheckboxColumn: true,
              columns: <GridColumn>[
                GridColumn(
                    columnName: 'id',
                    label: Container(
                        padding: EdgeInsets.all(16.0),
                        alignment: Alignment.center,
                        child: Text(
                          'ID',
                        ))),
                GridColumn(
                    columnName: 'name',
                    label: Container(
                        padding: EdgeInsets.all(8.0),
                        alignment: Alignment.center,
                        child: Text('Name'))),
                GridColumn(
                    columnName: 'designation',
                    label: Container(
                        padding: EdgeInsets.all(8.0),
                        alignment: Alignment.center,
                        child: Text(
                          'Designation',
                          overflow: TextOverflow.ellipsis,
                        ))),
                GridColumn(
                    columnName: 'salary',
                    label: Container(
                        padding: EdgeInsets.all(8.0),
                        alignment: Alignment.center,
                        child: Text('Salary'))),
              ],
            ),
          ),
        ],
      ),
    );
  }

  void downloadFile(String filename, List<int> bytes){
    html.AnchorElement anchorElement =  new html.AnchorElement();
    anchorElement.setAttribute("download", filename);
    anchorElement.href = 'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}';
    anchorElement.click();
  }

  Future<void> generateExcel() async {
    //Create a Excel document.

    //Creating a workbook.
    final Workbook workbook = Workbook();
    //Accessing via index
    final Worksheet sheet = workbook.worksheets[0];
    sheet.showGridlines = true;

    // Enable calculation for worksheet.
    sheet.enableSheetCalculations();
    sheet.getRangeByName('A1').setText('Ad');
    sheet.getRangeByName('B1').setText('Unvan');
    sheet.getRangeByName('C1').setText('Maaş');
    List<String> selectedAlphabets = alphabetCharacters.take(3).toList();
    for (var i = 0; i < employees.length; i++) {
      for(var j = 0; j < selectedAlphabets.length; j++){
        sheet.getRangeByName(selectedAlphabets[j] + (i+2).toString()).setText(getContentFromModel(employees[i], j));
      }
    }
    //Save and launch the excel.
    final List<int> bytes = workbook.saveAsStream();
    //Dispose the document.
    workbook.dispose();

    //Save and launch the file.
    downloadFile("excel_name.xlsx", bytes);
  }

  String getContentFromModel(Employee employee, int value) {
    String content = "";
    if (value == 0) {
      content = employee.name;
    } else if (value == 1) {
      content = employee.designation;
    } else {
      content = employee.salary.toString();
    }
    return content;
  }

  List<Employee> getEmployeeData() {
    return [
      Employee(10001, 'James', 'Project Lead', 20000),
      Employee(10002, 'Kathryn', 'Manager', 30000),
      Employee(10003, 'Lara', 'Developer', 15000),
      Employee(10004, 'Michael', 'Designer', 15000),
      Employee(10005, 'Martin', 'Developer', 15000),
      Employee(10006, 'Newberry', 'Developer', 15000),
      Employee(10007, 'Balnc', 'Developer', 15000),
      Employee(10008, 'Perry', 'Developer', 15000),
      Employee(10009, 'Gable', 'Developer', 15000),
      Employee(10010, 'Grimes', 'Developer', 15000)
    ];
  }
}

class Employee {
  /// Creates the employee class with required details.
  Employee(this.id, this.name, this.designation, this.salary);

  /// Id of an employee.
  final int id;

  /// Name of an employee.
  final String name;

  /// Designation of an employee.
  final String designation;

  /// Salary of an employee.
  final int salary;
}

/// An object to set the employee collection data source to the datagrid. This
/// is used to map the employee data to the datagrid widget.
class EmployeeDataSource extends DataGridSource {
  /// Creates the employee data source class with required details.
  EmployeeDataSource({required List<Employee> employeeData}) {
    _employeeData = employeeData
        .map<DataGridRow>((e) => DataGridRow(cells: [
      DataGridCell<int>(columnName: 'id', value: e.id),
      DataGridCell<String>(columnName: 'name', value: e.name),
      DataGridCell<String>(
          columnName: 'designation', value: e.designation),
      DataGridCell<int>(columnName: 'salary', value: e.salary),
    ]))
        .toList();
  }

  List<DataGridRow> _employeeData = [];

  @override
  List<DataGridRow> get rows => _employeeData;

  @override
  DataGridRowAdapter buildRow(DataGridRow row) {
    return DataGridRowAdapter(
        cells: row.getCells().map<Widget>((e) {
          return Container(
            alignment: Alignment.center,
            padding: EdgeInsets.all(8.0),
            child: Text(e.value.toString()),
          );
        }).toList());
  }
}