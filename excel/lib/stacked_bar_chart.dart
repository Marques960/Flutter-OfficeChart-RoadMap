// Imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:syncfusion_officechart/officechart.dart' as Chart;
import 'package:syncfusion_officechart/officechart.dart';

Future<void> stacked_bar_chart() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet5 = workbook.worksheets[0];

  sheet5.showGridlines = false;

  sheet5.getRangeByName('A1:C1').columnWidth = 30;

  sheet5.enableSheetCalculations();
  //header of the graphic
  sheet5.getRangeByName('A1').setText('Name');
  sheet5.getRangeByName('B1').setText('Salary');
  sheet5.getRangeByName('C1').setText('Working hr');
    //personalize
    sheet5.getRangeByName('A1:C1').cellStyle.hAlign = HAlignType.center;
    sheet5.getRangeByName('A1:C1').cellStyle.vAlign = VAlignType.center;
    sheet5.getRangeByName('A1:C1').cellStyle.fontSize = 18;
    sheet5.getRangeByName('A1:C1').cellStyle.bold = true;

  //name
  sheet5.getRangeByName('A2').setText('Ben');
  sheet5.getRangeByName('A3').setText('Mark');
  sheet5.getRangeByName('A4').setText('Sundar');
  sheet5.getRangeByName('A5').setText('Geo');
  sheet5.getRangeByName('A6').setText('Andrew');
  
  //salary
  sheet5.getRangeByName('B2').setNumber(1000);
  sheet5.getRangeByName('B3').setNumber(2000);
  sheet5.getRangeByName('B4').setNumber(2392);
  sheet5.getRangeByName('B5').setNumber(3211);
  sheet5.getRangeByName('B6').setNumber(4211);
  
  //working hr
  sheet5.getRangeByName('C2').setNumber(287);
  sheet5.getRangeByName('C3').setNumber(355);
  sheet5.getRangeByName('C4').setNumber(134);
  sheet5.getRangeByName('C5').setNumber(581);
  sheet5.getRangeByName('C6').setNumber(426);
    //personalization
    sheet5.getRangeByName('A2:C6').cellStyle.hAlign = HAlignType.center;
    sheet5.getRangeByName('A2:C6').cellStyle.vAlign = VAlignType.center;
    sheet5.getRangeByName('A2:C6').cellStyle.fontSize = 14;

  //create pie chart
  final charts = ChartCollection(sheet5);
  final chart1 = charts.add();
  //define as bar chart
  chart1.chartType = ExcelChartType.barStacked;

  //limit the data
  chart1.dataRange = sheet5.getRangeByName('A1:C6');
  sheet5.getRangeByName('A2:C6').rowHeight = 25;
  chart1.isSeriesInRows = false;

  //atribute title for the chart
  chart1.chartTitle = "Stacked Bar Chart";
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
  sheet5.charts = charts;

//----------------------------------------------------------------------------------
  //Save and launch the excel.
  final List<int> bytes = workbook.saveAsStream();

  print('Bytes length: ${bytes.length}');

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      //atribute some name do the excel
      ..setAttribute('download', 'Stacked_Bar_Chart.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Erro ao gerar os bytes do arquivo.');
  }
}
