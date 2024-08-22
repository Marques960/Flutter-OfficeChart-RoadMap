// Imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:syncfusion_officechart/officechart.dart' as Chart;

Future<void> line_chart() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet3 = workbook.worksheets[0];

  sheet3.showGridlines = false;

  sheet3.getRangeByName('A1:B1').columnWidth = 30;

  sheet3.enableSheetCalculations();
  //header of the graphic
  sheet3.getRangeByName('A1').setText('City Name');
  sheet3.getRangeByName('B1').setText('Temp in C');
    //personalize
    sheet3.getRangeByName('A1:B1').cellStyle.hAlign = HAlignType.center;
    sheet3.getRangeByName('A1:B1').cellStyle.vAlign = VAlignType.center;
    sheet3.getRangeByName('A1:B1').cellStyle.fontSize = 18;
    sheet3.getRangeByName('A1:B1').cellStyle.bold = true;
  //name of the cities
  sheet3.getRangeByName('A2').setText('Chennai');
  sheet3.getRangeByName('A3').setText('Mumbai');
  sheet3.getRangeByName('A4').setText('Delhi');
  sheet3.getRangeByName('A5').setText('Hyderabad');
  sheet3.getRangeByName('A6').setText('Kolkata');
    //personalization
    sheet3.getRangeByName('A2:A6').cellStyle.hAlign = HAlignType.center;
    sheet3.getRangeByName('A2:A6').cellStyle.vAlign = VAlignType.center;
    sheet3.getRangeByName('A2:A6').cellStyle.fontSize = 14;

  //temperature
  sheet3.getRangeByName('B2').setNumber(34);
  sheet3.getRangeByName('B3').setNumber(40);
  sheet3.getRangeByName('B4').setNumber(47);
  sheet3.getRangeByName('B5').setNumber(20);
  sheet3.getRangeByName('B6').setNumber(66);
    //center temperature
    sheet3.getRangeByName('B2:B6').cellStyle.hAlign = HAlignType.center;
    sheet3.getRangeByName('B2:B6').cellStyle.vAlign = VAlignType.center;
    sheet3.getRangeByName('B2:B6').cellStyle.fontSize = 14;

  //create line chart
  final chart_3 = Chart.ChartCollection(sheet3);
  final chart3 = chart_3.add();
  chart3.chartType = Chart.ExcelChartType.line;  
  //limit the data
  chart3.dataRange = sheet3.getRangeByName('A1:B6');
  sheet3.getRangeByName('A2:B6').rowHeight = 25;
  chart3.isSeriesInRows = false;

  //atribute title for the chart
  chart3.chartTitle = "Line Chart";
  chart3.chartTitleArea.bold;
  chart3.chartTitleArea.size = 20;

  for (int i = 0; i < chart3.series.count; i++) {
    final Chart.ChartSerie serie = chart3.series[i];
    serie.dataLabels.isValue = true;
  }

  //move chart
  chart3.topRow = 5;
  chart3.bottomRow = 25;
  chart3.rightColumn = 20;
  chart3.leftColumn = 6;
  chart3.plotArea;
  chart3.linePattern = Chart.ExcelChartLinePattern.solid;
  chart3.linePattern = Chart.ExcelChartLinePattern.longDashDotDot;
  sheet3.charts = chart_3;

//----------------------------------------------------------------------------------
  //Save and launch the excel.
  final List<int> bytes = workbook.saveAsStream();

  print('Bytes length: ${bytes.length}');

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      //atribute some name do the excel
      ..setAttribute('download', 'Line_chart.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Erro ao gerar os bytes do arquivo.');
  }
}
