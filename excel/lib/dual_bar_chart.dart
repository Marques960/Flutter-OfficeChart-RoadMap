// Imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:syncfusion_officechart/officechart.dart' as Chart;
import 'package:syncfusion_officechart/officechart.dart';

Future<void> dual_bar_chart() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet2 = workbook.worksheets[0];

  sheet2.showGridlines = false;

  sheet2.getRangeByName('A1:B1').columnWidth = 30;

  sheet2.enableSheetCalculations();
  //header of the graphic
  sheet2.getRangeByName('A1').setText('Items');
  sheet2.getRangeByName('B1').setText('Amount(in \$)');
  sheet2.getRangeByName('C1').setText('Count');
    //personalize
    sheet2.getRangeByName('A1:C1').cellStyle.hAlign = HAlignType.center;
    sheet2.getRangeByName('A1:C1').cellStyle.vAlign = VAlignType.center;
    sheet2.getRangeByName('A1:C1').cellStyle.fontSize = 16;
    sheet2.getRangeByName('A1:C1').cellStyle.bold = true;
  
  //categories
  
  sheet2.getRangeByName('A2').setText('Beverages');
  sheet2.getRangeByName('A3').setText('Condiments');
  sheet2.getRangeByName('A4').setText('Confections');
  sheet2.getRangeByName('A5').setText('Dairy Products');
  sheet2.getRangeByName('A6').setText('Grains / Cereals');
    //personalization
    sheet2.getRangeByName('A2:A6').cellStyle.hAlign = HAlignType.center;
    sheet2.getRangeByName('A2:A6').cellStyle.vAlign = VAlignType.center;
    sheet2.getRangeByName('A2:A6').cellStyle.fontSize = 14;

  //quantitites  
  sheet2.getRangeByName('B2').setNumber(2776);
  sheet2.getRangeByName('B3').setNumber(1077);
  sheet2.getRangeByName('B4').setNumber(2287);
  sheet2.getRangeByName('B5').setNumber(1368);
  sheet2.getRangeByName('B6').setNumber(3325);
  sheet2.getRangeByName('C2').setNumber(925);
  sheet2.getRangeByName('C3').setNumber(378);
  sheet2.getRangeByName('C4').setNumber(880);
  sheet2.getRangeByName('C5').setNumber(581);
  sheet2.getRangeByName('C6').setNumber(189);
    //center quantities
    sheet2.getRangeByName('B2:C6').cellStyle.hAlign = HAlignType.center;
    sheet2.getRangeByName('B2:C6').cellStyle.vAlign = VAlignType.center;
    sheet2.getRangeByName('B2:C6').cellStyle.fontSize = 14;

  //create pie chart
  final  charts = ChartCollection(sheet2);
  final  chart1 = charts.add();
  //define as bar chart
  chart1.chartType = ExcelChartType.columnStacked;

  //limit the data
  chart1.dataRange = sheet2.getRangeByName('A1:C6');
  sheet2.getRangeByName('A2:C6').rowHeight = 25;
  chart1.isSeriesInRows = false;

  //atribute title for the chart
  chart1.chartTitle = "Bar Chart";
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
  sheet2.charts = charts;

//----------------------------------------------------------------------------------
  //Save and launch the excel.
  final List<int> bytes = workbook.saveAsStream();

  print('Bytes length: ${bytes.length}');

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      //atribute some name do the excel
      ..setAttribute('download', 'Dual_bar_chart.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Erro ao gerar os bytes do arquivo.');
  }
}
