import 'dart:convert';
import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter/services.dart' show ByteData, rootBundle; // Para acessar os assets
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel to JSON Converter',
      theme: ThemeData(primarySwatch: Colors.blue),
      home: ExcelToJsonPage(),
    );
  }
}

class ExcelToJsonPage extends StatefulWidget {
  @override
  _ExcelToJsonPageState createState() => _ExcelToJsonPageState();
}

class _ExcelToJsonPageState extends State<ExcelToJsonPage> {
  String _status = "Pressione o botão para converter";

  void _convertExcelToJson() async {
    final ByteData data = await rootBundle.load('assets/xlsx_files/devs.xlsx');
    final List<int> bytes = data.buffer.asUint8List();
    var excel = Excel.decodeBytes(bytes);

    List<Map<String, dynamic>> jsonList = [];

    // Obtendo a primeira planilha
    var sheet = excel.tables[excel.tables.keys.first];

    if (sheet == null || sheet.rows.isEmpty) {
      print("Erro: A planilha está vazia ou não foi encontrada.");
      return;
    }

    // Pegando a primeira linha como cabeçalhos (atributos)
    List<String> headers = [];
    for (var cell in sheet.rows.first) {
      headers.add(cell?.value.toString() ?? ""); // Pegando apenas o valor da célula
    }

    // Iterar a partir da segunda linha (índice 1) para extrair os dados
    for (var i = 1; i < sheet.rows.length; i++) {
      var row = sheet.rows[i];
      Map<String, dynamic> jsonRow = {};

      for (var j = 0; j < row.length; j++) {
        if (j < headers.length) {
          jsonRow[headers[j]] = row[j]?.value.toString() ?? ""; // Pegando apenas o valor
        }
      }

      jsonList.add(jsonRow);
    }

    // Convertendo para JSON
    String jsonString = jsonEncode(jsonList);

    // Diretório de Downloads no armazenamento externo
    final downloadDirPath = Directory('/storage/emulated/0/Download');

    if (!await downloadDirPath.exists()) {
      await downloadDirPath.create(recursive: true);
    }

    // Caminho do arquivo JSON
    final jsonFilePath = '${downloadDirPath.path}/devs1.json';

    // Salvar o arquivo JSON
    final jsonFile = File(jsonFilePath);
    await jsonFile.writeAsString(jsonString);

    print('Arquivo JSON salvo em: $jsonFilePath');
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("Converter Excel para JSON")),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            ElevatedButton(onPressed: _convertExcelToJson, child: Text("Converter Excel para JSON")),
            SizedBox(height: 20),
            Text(_status, textAlign: TextAlign.center),
          ],
        ),
      ),
    );
  }
}
