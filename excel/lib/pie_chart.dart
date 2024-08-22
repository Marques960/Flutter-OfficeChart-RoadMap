//imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:syncfusion_officechart/officechart.dart';


Future<void> pie_chart() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet_1 = workbook.worksheets[0];
  sheet_1.showGridlines = false;

  sheet_1.enableSheetCalculations();

  sheet_1.getRangeByName('A11').setText('Venue');
  sheet_1.getRangeByName('A12').setText('Seating & Decor');
  sheet_1.getRangeByName('A13').setText('Technical Team');
  sheet_1.getRangeByName('A14').setText('performers');
  sheet_1.getRangeByName('A15').setText('performer\'s Transport');
  sheet_1.getRangeByName('B11:B15').numberFormat = '\$#,##0_)';
  sheet_1.getRangeByName('B11').setNumber(17500);
  sheet_1.getRangeByName('B12').setNumber(1828);
  sheet_1.getRangeByName('B13').setNumber(800);
  sheet_1.getRangeByName('B14').setNumber(14000);
  sheet_1.getRangeByName('B15').setNumber(2600);

  // Add the chart to the worksheet_1.
  final chart = ChartCollection(sheet_1);
  final chart_1 = chart.add();

  // Set chart properties.
  chart_1.chartType = ExcelChartType.pie;
  chart_1.dataRange = sheet_1.getRangeByName('A11:B15');
  chart_1.isSeriesInRows = false;
  //chart_1.topRow = 1;
  //chart_1.leftColumn = 4;
  //chart_1.bottomRow = 10;
  //chart_1.rightColumn = 9;

  // Save and launch the Excel file
  final List<int> bytes = workbook.saveAsStream();

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      ..setAttribute('download', 'pie_chart.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Error generating the file bytes.');
  }

  workbook.dispose();
}
