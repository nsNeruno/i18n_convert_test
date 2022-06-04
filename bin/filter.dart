import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';

void transfer(Map<String, dynamic> ref, Excel src, Excel dst, String sheetName,) {
  final sheet = src.sheets[sheetName];
  dst.copy(dst.getDefaultSheet() ?? sheetName, sheetName,);
  final dstSheet = dst.sheets[sheetName]!;
  if (sheet != null) {
    final rows = Map<String, String>.fromEntries(
      sheet.rows.where(
        (List<Data?> row,) => row[0]?.toString().isNotEmpty == true,
      ).map(
        (List<Data?> row,) => MapEntry<String, String>(row[0]!.value.toString(), row[1]?.value.toString() ?? "",),
      ),
    );
    ref.forEach(
      (key, value) {
        dstSheet.appendRow(
          [key, rows[key] ?? '',],
        );
      },
    );
  }
}

void main(List<String> _,) async {
  final String json = await File("en_US.json",).readAsString();
  final Map<String, dynamic> parsed = jsonDecode(json,);
  final excel = Excel.decodeBytes(
    File("Translation consolidated.xlsx",).readAsBytesSync(),
  );
  final targetExcel = Excel.createExcel();
  final languages = [
    'fr_FR',
    'es_ES',
    'ml_MY',
  ];
  for (String language in languages) {
    transfer(parsed, excel, targetExcel, language,);
  }
  final dstBytes = targetExcel.save(fileName: 'Consolidated.xlsx',);
  if (dstBytes != null) {
    await File('Consolidated.xlsx',).writeAsBytes(dstBytes,);
  } else {
    print("Missing bytes",);
  }
}