import 'dart:collection';
import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';

const List<String> languageKeys = ["en_US", "tl_PH", "id_ID",];

Map<String, String?> readOldSheet(Sheet? sheet, [int valueIndex = 1,]) {
  var map = <String, String?>{};
  sheet?.rows.forEach(
    (List<Data?> dataRow) {
      final String? key = dataRow[0]?.value.toString().replaceAll("&", "N",);
      final String? value = dataRow[valueIndex]?.value;
      if (key != null) {
        map[key] = value;
      }
    },
  );
  return map;
}

void _writeSheet(Sheet? sheet, String fileName, {
  int valueIndex = 1, bool createDartFile = false,
}) {
  if (sheet == null) {
    return;
  }
  var map = readOldSheet(sheet, valueIndex,);
  File(fileName,).writeAsString(
    JsonEncoder.withIndent("\t",).convert(
      map,
    ),
  );
  if (createDartFile) {
    final writer = File("localization_keys.dart",).openWrite();
    writer.writeln(
      "class Translated {",
    );
    writer.writeln();
    writer.writeln("\tTranslated._();",);
    writer.writeln();
    map.keys.forEach(
      (String key) {
        writer.writeln(
          "\tstatic const $key = \"$key\";",
        );
      },
    );
    writer.write("}",);
  }
}

Map<String, Map<String, String?>> readPatchExcelFile(Excel excel, String patchSheetName,) {
  final sheet = excel[patchSheetName];
  final result = <String, Map<String, String?>>{};
  final List<List<Data?>> rows = sheet.rows;
  for (int i = 1; i < rows.length; i++) {
    final List<Data?> row = rows[i];
    final String key = row[0]?.value.toString() ?? '';
    for (int j = 0; j < languageKeys.length; j++) {
      final lang = languageKeys[j];
      result[lang] ??= {};
      result[lang]?[key] = row[j + 1]?.value.toString();
    }
  }
  return result;
}

Map<String, String?> mergeSorted(Map<String, String?>? a, Map<String, String?>? b,) {
  if (b != null) {
    a?.addAll(b,);
  }
  return SplayTreeMap.from(a ?? {},);
}

void main(List<String> args) async {
  var bytes = await File("DigLog_mPOS_i18n - Draft.xlsx",).readAsBytes();
  var excel = Excel.decodeBytes(bytes,);
  final Map<String, Map<String, String?>> master = {};
  for (String key in languageKeys) {
    final Sheet? sheet = excel.sheets[key];
    master[key] = readOldSheet(sheet,);
  }

  // var sheet = excel.tables["en_US"];
  // _writeSheet(sheet, "en_US.json", createDartFile: true,);
  // bytes = await File("Mobile app - Bahasa-Tagalog version.xlsx",).readAsBytes();
  // excel = Excel.decodeBytes(bytes,);
  // _writeSheet(excel.tables["Bahasa"], "id_ID.json", valueIndex: 2,);
  // _writeSheet(excel.tables["Sheet2"], "tl_PH.json", valueIndex: 2,);
  bytes = await File("DigLog_mPOS_i18n - translated.xlsx",).readAsBytes();
  excel = Excel.decodeBytes(bytes,);
  final Map<String, Map<String, String?>> result = readPatchExcelFile(
    excel, "New Patches",
  );
  for (String key in languageKeys) {
    master[key] = mergeSorted(master[key], result[key],);
  }
  // master.entries.forEach(
  //   (entry) {
  //     print(entry.key,);
  //     entry.value.forEach(
  //       (key, value) {
  //         print("[$key]: $value",);
  //       },
  //     );
  //   },
  // );
  final masterExcel = Excel.createExcel();
  final defaultSheetName = masterExcel.getDefaultSheet();
  if (defaultSheetName != null) {
    if (defaultSheetName != languageKeys[0]) {
      masterExcel.rename(defaultSheetName, languageKeys[0],);
      for (int i = 1; i < languageKeys.length; i++) {
        masterExcel.copy(languageKeys[0], languageKeys[i],);
      }
    }
  }
  master.entries.forEach(
    (entry) {
      final Sheet? sheet = masterExcel.sheets[entry.key];
      final entries = entry.value.entries.toList();
      for (int i = 0; i < entries.length; i++) {
        sheet?.insertRowIterables(
          [entries[i].key, entries[i].value,], i,
        );
      }
    },
  );
  final masterFile = File("DigLog_mPOS_i18n - Master.xlsx",);
  final saveResult = masterExcel.save(fileName: "DigLog_mPOS_i18n - Master.xlsx",);
  if (saveResult != null) {
    masterFile.writeAsBytesSync(saveResult,);
  }
}