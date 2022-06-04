import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';

Future<void> writeJson(String filename, Excel excel,) async {
  final jsonContent = await File("$filename.json").readAsString();
  final Map<String, dynamic> decoded = jsonDecode(jsonContent,);
  final sheet = excel.sheets[filename];
  final entries = decoded.entries.toList();
  for (int i = 0; i < entries.length; i++) {
    sheet?.insertRowIterables(
      [entries[i].key, entries[i].value.toString()], i,
    );
  }
}

void main(List<String> _) async {
  final masterExcel = Excel.createExcel();
  final masterFile = File("DigLog_mPOS_i18n - Master.xlsx",);
  final filenames = ["en_US", "id_ID", "tl_PH"];
  masterExcel.rename(masterExcel.getDefaultSheet() ?? "", filenames[0],);
  masterExcel.copy(filenames[0], filenames[1]);
  masterExcel.copy(filenames[0], filenames[2]);
  for (String filename in filenames) {
    await writeJson(filename, masterExcel,);
  }
  final saveResult = masterExcel.save(fileName: "DigLog_mPOS_i18n - Master.xlsx",);
  if (saveResult != null) {
    masterFile.writeAsBytesSync(saveResult,);
  }
}