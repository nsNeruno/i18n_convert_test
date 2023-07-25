import 'dart:collection';
import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';

const Set<String> languageKeys = {
  "en_US", // English
  "tl_PH", // Tagalog
  "id_ID", // Bahasa
  "fr_FR", // French
  "es_ES", // Espanol
  "ms_MY", // Melayu
  "zh", // Chinese
  "vi_VN", // Viet
  "ja_JP", // Japanese
  "hi", // Hindi
  "bn", // Bengali
  "te", // Telugu
  "ar", // Arabic
  "my", // Burmese (Myanmar)
  "sq_AL", // Albanian
  "bg_BG", // Bulgarian
  "hr_HR", // Croatian
  "cz_CZ", // Czech
  "da_DK", // Danish (Denmark)
  "nl_NL", // Dutch (Netherlands)
  "et_EE", // Estonian
  "fi_FI", // Finnish (Finland)
  "ka_GE", // Georgian
  "de_DE", // German
  "el_GR", // Greek
  "he", // Hebrew
  "hu_HU", // Hungarian
  "it_IT", // Italian
  "km_KH", // Khmer (Cambodian)
  "ko_KR", // Korean
  "lo_LA", // Laotian (Laos)
  "lv_LV", // Latvian
  "lt_LT", // Lithuanian
  "mk_MK", // Macedonian
  "mn_MN", // Mongolian
  "no_NO", // Norwegian
  "pl_PL", // Polish
  "pt_PT", // Portuguese
  "ro", // Romanian / Moldovan
  "sr", // Serbian
  "sv_SE", // Swedish
  "th_TH", // Thai
  "tr_TR", // Turkish
  "ur", // Urdu
  "uz_UZ", // Uzbek
};

Future<void> writeJsonFile(Sheet? sheet, String key, {
  bool createDartFile = false,
}) async {
  final jsonFile = File("results/$key.json",);
  if (!await jsonFile.exists()) {
    await jsonFile.create(recursive: true,);
  }
  var map = <String, String?>{};
  sheet?.rows.forEach(
    (List<Data?> dataRow) {
      final String? key = dataRow[0]?.value.toString().replaceAll("&", "N",);
      final String? value = dataRow[1]?.value.toString();
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
    final String key = languageKeys.elementAt(i,);
    var sheet = excel.sheets[key];
    writeJsonFile(sheet, key, createDartFile: i == 0,);
  }
  Process.run("open", ["results",],);
}