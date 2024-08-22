// Imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:syncfusion_officechart/officechart.dart' as Chart;
import 'package:syncfusion_officechart/officechart.dart';

Future<void> stacked_chart() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet4 = workbook.worksheets[0];

  sheet4.showGridlines = false;

  sheet4.getRangeByName('A1:C1').columnWidth = 30;

  sheet4.enableSheetCalculations();
  //header of the graphic
  sheet4.getRangeByName('A1').setText('Items');
  sheet4.getRangeByName('B1').setText('Amount(in \$)');
  sheet4.getRangeByName('C1').setText('Count');
    //personalize
    sheet4.getRangeByName('A1:C1').cellStyle.hAlign = HAlignType.center;
    sheet4.getRangeByName('A1:C1').cellStyle.vAlign = VAlignType.center;
    sheet4.getRangeByName('A1:C1').cellStyle.fontSize = 18;
    sheet4.getRangeByName('A1:C1').cellStyle.bold = true;
  
  //items
  sheet4.getRangeByName('A2').setText('Beverages');
  sheet4.getRangeByName('A3').setText('Condiments');
  sheet4.getRangeByName('A4').setText('Confections');
  sheet4.getRangeByName('A5').setText('Dairy Products');
  sheet4.getRangeByName('A6').setText('Grains / Cereals');

  //amount
  sheet4.getRangeByName('B2').setNumber(2776);
  sheet4.getRangeByName('B3').setNumber(1077);
  sheet4.getRangeByName('B4').setNumber(2287);
  sheet4.getRangeByName('B5').setNumber(1368);
  sheet4.getRangeByName('B6').setNumber(3325);

  //count
  sheet4.getRangeByName('C2').setNumber(925);
  sheet4.getRangeByName('C3').setNumber(378);
  sheet4.getRangeByName('C4').setNumber(880);
  sheet4.getRangeByName('C5').setNumber(581);
  sheet4.getRangeByName('C6').setNumber(189);
    //personalization
    sheet4.getRangeByName('A2:C6').cellStyle.hAlign = HAlignType.center;
    sheet4.getRangeByName('A2:C6').cellStyle.vAlign = VAlignType.center;
    sheet4.getRangeByName('A2:C6').cellStyle.fontSize = 14;

  //create pie chart
  final charts = ChartCollection(sheet4);
  final chart1 = charts.add();
  //define as bar chart
  chart1.chartType = ExcelChartType.columnStacked;

  //limit the data
  chart1.dataRange = sheet4.getRangeByName('A1:C6');
  sheet4.getRangeByName('A2:C6').rowHeight = 25;
  chart1.isSeriesInRows = false;

  //atribute title for the chart
  chart1.chartTitle = "Stacked Chart";
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
  sheet4.charts = charts;

//----------------------------------------------------------------------------------
  //Save and launch the excel.
  final List<int> bytes = workbook.saveAsStream();

  print('Bytes length: ${bytes.length}');

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      //atribute some name do the excel
      ..setAttribute('download', 'Stacked_bar.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Erro ao gerar os bytes do arquivo.');
  }
}
