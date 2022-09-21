import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({Key? key}) : super(key: key);

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Excel',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: const MyHomePage(),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({Key? key}) : super(key: key);

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  String? excelPath = "Upload excel";
  List<String> items = [];
  List<String> values = [];
  List<String> listOfOutPutForItems = [];
  List<String> listOfOutPutForValues = [];

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(),
      body: Scrollbar(
        child: SingleChildScrollView(
          child: Column(
            children: [
              TextButton(
                  onPressed: () async {
                    await uploadExcelSheet();
                  },
                  child: Text(excelPath!)),
              const SizedBox(height: 20),
              TextFormField(onChanged: (text) {
                text.toLowerCase();
                listOfOutPutForItems = [];
                listOfOutPutForValues = [];
                int indexOfElement = -1;
                if (text != "") {
                  for (var element in items) {
                    if (element.toLowerCase().contains(text.toLowerCase())) {
                      indexOfElement = items.indexOf(element);
                      listOfOutPutForValues.add(values[indexOfElement]);
                      listOfOutPutForItems.add(element);
                    }
                  }
                } else {
                  listOfOutPutForItems.clear();
                  listOfOutPutForValues.clear();
                }
                setState(() {});
              }),
              listOfOutPutForValues.isNotEmpty
                  ? Column(
                      children: listOfOutPutForValues.map((e) {
                        return Row(
                          children: [
                            SizedBox(child: SelectableText(itemText(e)!)),
                            const SizedBox(width: 20),
                            SizedBox(child: SelectableText(e)),
                          ],
                        );
                      }).toList(),
                    )
                  : const SizedBox(),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> uploadExcelSheet() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles();
    var bytes = File(result!.paths.first!).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    setState(() {
      excelPath = result.paths.first;
    });

    for (var table in excel.tables.keys) {
      int cols = excel.tables[table]!.maxCols - 1;
      for (var row in excel.tables[table]!.rows) {
        for (int j = 0; j < cols; j++) {
          if (row.length > j + 1) {
            items.add(row[j].toString());
            values.add(row[j + 1].toString());
            j++;
          }
        }
      }
    }
    for (int i = 0; i < items.length; i++) {
      items[i] = items[i].toLowerCase();
      values[i] = values[i].toLowerCase();
    }
  }

  String? itemText(String? text) {
    int indexOfValue = -1;
    indexOfValue = listOfOutPutForValues.indexOf(text!);
    String? itemString = listOfOutPutForItems[indexOfValue];
    return itemString;
  }
}
