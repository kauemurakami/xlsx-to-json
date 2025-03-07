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

    // Assumindo que você está pegando a primeira planilha
    var sheet = excel.tables[excel.tables.keys.first];

    // Iterar pelas linhas e colunas para extrair os valores
    for (var row in sheet!.rows) {
      Map<String, dynamic> jsonRow = {};
      for (var i = 0; i < row.length; i++) {
        // Verifique se a célula não é nula
        if (row[i] != null) {
          jsonRow['col$i'] = row[i].toString(); // Converte tudo para String
        }
      }
      jsonList.add(jsonRow);
    }

    // Agora você pode converter para JSON sem problemas
    String jsonString = jsonEncode(jsonList);

    // Acessando o diretório de Downloads diretamente no armazenamento externo
    final directory = await getExternalStorageDirectory();

    // Caminho do diretório de Downloads diretamente (não no diretório do app)
    final downloadDirPath = Directory('/storage/emulated/0/Download');

    // Verifique se o diretório de Downloads existe
    if (!await downloadDirPath.exists()) {
      await downloadDirPath.create(recursive: true);
    }

    // Caminho do arquivo JSON na pasta Downloads
    final jsonFilePath = '${downloadDirPath.path}/devs.json';

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
