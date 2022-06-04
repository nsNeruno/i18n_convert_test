import 'dart:io';

import 'package:excel/excel.dart';

Future<List<String>?> getMasterKeys() async {
  final data = await File("DigLog_mPOS_i18n - Master.xlsx").readAsBytes();
  final excel = Excel.decodeBytes(data);
  final sheet = excel.sheets['en_US'];
  return sheet?.rows.map<String?>(
    (List<Data?> row,) => row[0]?.value.toString(),
  ).whereType<String>().toList();
}
void main(List<String> _,) async {

}