//ignorances
// ignore_for_file: non_constant_identifier_names 
// ignore_for_file: unused_local_variable
// ignore_for_file: avoid_print

//imports
import 'dart:typed_data';
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart';

Future<void> listed_sheet() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet = workbook.worksheets[0];
  sheet.showGridlines = false;

  List<int> lista_ids_tasks = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  List<String> nomes_funcs_excel = ['teste1', 'teste2', 'teste3', 'teste4', 'teste5', 'teste6', 'teste7', 'teste8', 'teste9', 'teste10'];
  List<String> vmp_excel = ['0.11', '0.22', '0.33', '0.44', '0.55', '0.66', '0.88', '0.99', '1.11', '2.11'];
  List<String> percent_erro_excel = ['0.28', '0.32', '0.45', '0.28', '0.44', '0.95', '1.5', '4.28', '0.53', '2.10'];
  

  sheet.enableSheetCalculations();

    //Set data in the worksheet.
    sheet.getRangeByName('A1').columnWidth = 5;
    sheet.getRangeByName('B1:F1').columnWidth = 20;

    sheet.getRangeByName('A1:F1').cellStyle.backColor = '#333F4F';
    sheet.getRangeByName('A1:F1').merge();
    sheet.getRangeByName('B4:D6').merge();

    sheet.getRangeByName('B4').setText('Registo Avaliação');
    sheet.getRangeByName('B4').cellStyle.fontSize = 30;

    sheet.getRangeByName('B8').setText('Realizado por:');
    sheet.getRangeByName('B8').cellStyle.fontSize = 14;
    sheet.getRangeByName('B8').cellStyle.bold = true;

    sheet.getRangeByName('B9').setText('Rafael Marques');
    sheet.getRangeByName('B9').cellStyle.fontSize = 12;

    sheet.getRangeByName('B10').setText('Portugal');
    sheet.getRangeByName('B10').cellStyle.fontSize = 11;

    sheet.getRangeByName('B11').setText('Contactos:');
    sheet.getRangeByName('B11').cellStyle.fontSize = 14;
    sheet.getRangeByName('B11').cellStyle.bold = true;

    sheet.getRangeByName('B12').setNumber(935485948);
    sheet.getRangeByName('B12').cellStyle.fontSize = 11;
    sheet.getRangeByName('B12').cellStyle.hAlign = HAlignType.left;


    sheet.getRangeByName('F8').setText('Numero Task (#)');
    sheet.getRangeByName('F8').cellStyle.fontSize = 14;
    sheet.getRangeByName('F8').cellStyle.bold = true;
    sheet.getRangeByName('F8').cellStyle.hAlign = HAlignType.right;

    sheet.getRangeByName('F9').setNumber(1000000014);
    sheet.getRangeByName('F9').cellStyle.fontSize = 10;
    sheet.getRangeByName('F9').cellStyle.hAlign = HAlignType.right;

    sheet.getRangeByName('F10').setText('Data');
    sheet.getRangeByName('F10').cellStyle.fontSize = 14;
    sheet.getRangeByName('F10').cellStyle.bold = true;
    sheet.getRangeByName('F10').cellStyle.hAlign = HAlignType.right;

    sheet.getRangeByName('F11').dateTime = DateTime.now();
    sheet.getRangeByName('F11').numberFormat ='[\$-x-sysdate]dddd, mmmm dd, yyyy';
    sheet.getRangeByName('F11').cellStyle.fontSize = 9;
    sheet.getRangeByName('F11').cellStyle.hAlign = HAlignType.right;

    final Range range6 = sheet.getRangeByName('B15:G15');
    range6.cellStyle.fontSize = 10;
    range6.cellStyle.bold = true;

    //headers
    //-num task
    sheet.getRangeByName('B15').setText('Núm Task');
    sheet.getRangeByName('B15').cellStyle.fontSize = 14;
    sheet.getRangeByName('B15').cellStyle.backColor = "#ACB9CA";
    sheet.getRangeByName('B15').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('B15').cellStyle.vAlign = VAlignType.center;
    

    //nome func
    sheet.getRangeByName('C15').setText('Nome Func.');
    sheet.getRangeByName('C15').cellStyle.fontSize = 14;
    sheet.getRangeByName('C15').cellStyle.backColor = "#ACB9CA";
    sheet.getRangeByName('C15').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('C15').cellStyle.vAlign = VAlignType.center;
    
    //vmp
    sheet.getRangeByName('D15').setText('V.M.P');
    sheet.getRangeByName('D15').cellStyle.fontSize = 14;
    sheet.getRangeByName('D15').cellStyle.backColor = "#ACB9CA";
    sheet.getRangeByName('D15').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('D15').cellStyle.vAlign = VAlignType.center;

    //-valor
    sheet.getRangeByName('E15').setText('Valor Atb.');
    sheet.getRangeByName('E15').cellStyle.fontSize = 14;
    sheet.getRangeByName('E15').cellStyle.backColor = "#ACB9CA";
    sheet.getRangeByName('E15').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('E15').cellStyle.vAlign = VAlignType.center;

    //-valor
    sheet.getRangeByName('F15').setText('Valor Atb.');
    sheet.getRangeByName('F15').cellStyle.fontSize = 14;
    sheet.getRangeByName('F15').cellStyle.backColor = "#ACB9CA";
    sheet.getRangeByName('F15').cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName('F15').cellStyle.vAlign = VAlignType.center;

    for (int i = 0; i < lista_ids_tasks.length; i++) {
      int rowIndex = 16 + i;
      //numero task
      final range1 = sheet.getRangeByName('B$rowIndex');
      range1.setNumber(lista_ids_tasks[i].toDouble());
      range1.cellStyle.fontSize = 12;
      range1.cellStyle.hAlign = HAlignType.center;
      range1.cellStyle.vAlign = VAlignType.center;
      
      //nome do funcionario
      final range2 = sheet.getRangeByName('C$rowIndex');
      range2.setText(nomes_funcs_excel[i].toString());
      range2.cellStyle.fontSize = 12;
      range2.cellStyle.hAlign = HAlignType.center;
      range2.cellStyle.vAlign = VAlignType.center;
      
      //vmp
      final range3 = sheet.getRangeByName('D$rowIndex');
      range3.setText(vmp_excel[i].toString());
      range3.cellStyle.fontSize = 12;
      range3.cellStyle.hAlign = HAlignType.center;
      range3.cellStyle.vAlign = VAlignType.center;
      
      //peecent erro
      final range4 = sheet.getRangeByName('E$rowIndex');
      range4.setText(percent_erro_excel[i].toString());
      range4.cellStyle.fontSize = 12;
      range4.cellStyle.hAlign = HAlignType.center;
      range4.cellStyle.vAlign = VAlignType.center;
    }

    sheet.getRangeByIndex(46, 1).text ='DecodeFY';
    sheet.getRangeByIndex(46, 1).cellStyle.fontSize = 14;

    final Range range9 = sheet.getRangeByName('A46:F47');
    range9.cellStyle.backColor = '#ACB9CA';
    range9.merge();
    range9.cellStyle.hAlign = HAlignType.center;
    range9.cellStyle.vAlign = VAlignType.center;

    //Save and launch the excel.
  final List<int> bytes = workbook.saveAsStream();
  
  print('Bytes length: ${bytes.length}'); 

  if (bytes != null && bytes.isNotEmpty) {
    final blob = html.Blob([Uint8List.fromList(bytes)]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      ..setAttribute('download', 'listed sheet.xlsx')
      ..click();

    html.Url.revokeObjectUrl(url);
  } else {
    print('Erro ao gerar os bytes do arquivo.');
  }
}