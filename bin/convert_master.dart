import 'dart:collection';
import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';

const List<String> languageKeys = [
  "en_US",
  "tl_PH",
  "id_ID",
  "fr_FR",
  "es_ES",
  "ms_MY",
  "zh",
  "vi_VN",
  "ja_JP",
  "hi",
  "bn",
  "te",
  "my",
  "ar",
  ];

Future<void> writeJsonFile(Sheet? sheet, String key, {bool createDartFile = false,}) async {
  final jsonFile = File("results/$key.json",);
  if (!await jsonFile.exists()) {
    await jsonFile.create(recursive: true,);
  }
  var map = <String, String?>{};
  sheet?.rows.forEach(
    (List<Data?> dataRow) {
      final String? key = dataRow[0]?.value.toString().replaceAll("&", "N",);
      final String? value = dataRow[1]?.value;
      if (key != null) {
        map[key] = value;
      }
    },
  );
  map = SplayTreeMap.from(map,);
  jsonFile.writeAsString(
    JsonEncoder.withIndent("\t",).convert(
      map,
    ),
  );
  if (createDartFile) {
    final writer = File("results/localization_keys.dart",).openWrite();
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
    await writer.flush();
    await writer.close();
  }
}

void main(List<String> args,) async {
  var bytes = await File("DizLog_mPOS_i18n - Master.xlsx",).readAsBytes();
  var excel = Excel.decodeBytes(bytes,);
  for (int i = 0; i < languageKeys.length; i++) {
    final String key = languageKeys[i];
    var sheet = excel.sheets[key];
    writeJsonFile(sheet, key, createDartFile: i == 0,);
  }
  Process.run("open", ["results",],);
}