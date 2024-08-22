// Imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:syncfusion_officechart/officechart.dart' as Chart;
import 'package:syncfusion_officechart/officechart.dart';

Future<void> stacked_line_chart() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet6 = workbook.worksheets[0];

  sheet6.showGridlines = false;

  sheet6.getRangeByName('A1:C1').columnWidth = 30;

  sheet6.enableSheetCalculations();
  //header of the graphic
  sheet6.getRangeByName('A1').setText('City Name');
  sheet6.getRangeByName('B1').setText('Temp in C');
  sheet6.getRangeByName('C1').setText('Temp in F');
    //personalize
    sheet6.getRangeByName('A1:C1').cellStyle.hAlign = HAlignType.center;
    sheet6.getRangeByName('A1:C1').cellStyle.vAlign = VAlignType.center;
    sheet6.getRangeByName('A1:C1').cellStyle.fontSize = 18;
    sheet6.getRangeByName('A1:C1').cellStyle.bold = true;

  sheet6.getRangeByName('A2').setText('Chennai');
  sheet6.getRangeByName('A3').setText('Mumbai');
  sheet6.getRangeByName('A4').setText('Delhi');
  sheet6.getRangeByName('A5').setText('Hyderabad');
  sheet6.getRangeByName('A6').setText('Kolkata');
  
  sheet6.getRangeByName('B2').setNumber(34);
  sheet6.getRangeByName('B3').setNumber(40);
  sheet6.getRangeByName('B4').setNumber(47);
  sheet6.getRangeByName('B5').setNumber(20);
  sheet6.getRangeByName('B6').setNumber(66);
  
  sheet6.getRangeByName('C2').setNumber(93);
  sheet6.getRangeByName('C3').setNumber(104);
  sheet6.getRangeByName('C4').setNumber(120);
  sheet6.getRangeByName('C5').setNumber(80);
  sheet6.getRangeByName('C6').setNumber(140);
    //personalization
    sheet6.getRangeByName('A2:C6').cellStyle.hAlign = HAlignType.center;
    sheet6.getRangeByName('A2:C6').cellStyle.vAlign = VAlignType.center;
    sheet6.getRangeByName('A2:C6').cellStyle.fontSize = 14;

  //create pie chart
  final charts = ChartCollection(sheet6);
  final chart1 = charts.add();
  //define as bar chart
  chart1.chartType = ExcelChartType.lineStacked;

  //limit the data
  chart1.dataRange = sheet6.getRangeByName('A1:C6');
  sheet6.getRangeByName('A2:C6').rowHeight = 25;
  chart1.isSeriesInRows = false;

  //atribute title for the chart
  chart1.chartTitle = "Stacked Line Chart";
  chart1.chartTitleArea.bold;
  chart1.chartTitleArea.size = 20;

  for (int i = 0; i < chart1.series.count; i++) {
    final Chart.ChartSerie serie = chart1.series[i];
    serie.dataLabels.isValue = true;
  }

  //move chart
  chart1.topRow = 5;
  chart1.bottomRow = 25;
  chart1.rightColumn = 20;
  chart1.leftColumn = 6;
  chart1.plotArea;
  chart1.linePattern = Chart.ExcelChartLinePattern.solid;
  chart1.linePattern = Chart.ExcelChartLinePattern.longDashDotDot;
  sheet6.charts = charts;

//----------------------------------------------------------------------------------
  //Save and launch the excel.
  final List<int> bytes = workbook.saveAsStream();

  print('Bytes length: ${bytes.length}');

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      //atribute some name do the excel
      ..setAttribute('download', 'Stacked_Line_Chart.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Erro ao gerar os bytes do arquivo.');
  }
}
