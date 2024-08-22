// Imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:syncfusion_officechart/officechart.dart' as Chart;

Future<void> pie_chart() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet = workbook.worksheets[0];

  sheet.showGridlines = false;

  sheet.getRangeByName('A1:B1').columnWidth = 30;

  sheet.enableSheetCalculations();
  //header of the graphic
  sheet.getRangeByName('A1').setText('Category');
  sheet.getRangeByName('B1').setText('Amount');
    //personalize
    sheet.getRangeByName('A1:B1').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('A1:B1').cellStyle.vAlign = VAlignType.center;
    sheet.getRangeByName('A1:B1').cellStyle.fontSize = 18;
    sheet.getRangeByName('A1:B1').cellStyle.bold = true;

  //personalization
  sheet.getRangeByName('A2').setText('Venue');
  sheet.getRangeByName('A3').setText('Seating & Technical');
  sheet.getRangeByName('A4').setText('Technical');
  sheet.getRangeByName('A5').setText('Performers');
  sheet.getRangeByName('A6').setText('Performers Transport');
    //personalization
    sheet.getRangeByName('A2:A6').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('A2:A6').cellStyle.vAlign = VAlignType.center;
    sheet.getRangeByName('A2:A6').cellStyle.fontSize = 14;

  //quantities
  sheet.getRangeByName('B2').setNumber(17500);
  sheet.getRangeByName('B3').setNumber(1828);
  sheet.getRangeByName('B4').setNumber(800);
  sheet.getRangeByName('B5').setNumber(14000);
  sheet.getRangeByName('B6').setNumber(2600);
    //center quantities
    sheet.getRangeByName('B2:B6').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('B2:B6').cellStyle.vAlign = VAlignType.center;
    sheet.getRangeByName('B2:B6').cellStyle.fontSize = 14;

  //create pie chart
  final chart_0 = Chart.ChartCollection(sheet);
  final chart0 = chart_0.add();
  //define chart as circular
  chart0.chartType = Chart.ExcelChartType.pie;
  //limit the data 
  chart0.dataRange = sheet.getRangeByName('A1:B6');
  sheet.getRangeByName('A2:B6').rowHeight = 25;
  chart0.isSeriesInRows = false;

  //atribute title for the chart
  chart0.chartTitle = "Pie Chart";
  chart0.chartTitleArea.bold;
  chart0.chartTitleArea.size = 20;

  for (int i = 0; i < chart0.series.count; i++) {
    final Chart.ChartSerie serie = chart0.series[i];
    serie.dataLabels.isValue = true;
  }

  //move chart
  chart0.topRow = 5;
  chart0.bottomRow = 25;
  chart0.rightColumn = 20;
  chart0.leftColumn = 6;
  chart0.plotArea;
  chart0.linePattern = Chart.ExcelChartLinePattern.solid;
  chart0.linePattern = Chart.ExcelChartLinePattern.longDashDotDot;
  sheet.charts = chart_0;

//----------------------------------------------------------------------------------
  //Save and launch the excel.
  final List<int> bytes = workbook.saveAsStream();

  print('Bytes length: ${bytes.length}');

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
    //atribute some name do the excel
      ..setAttribute('download', 'Excel_Comp_Funcs.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Erro ao gerar os bytes do arquivo.');
  }
}
