// ignore_for_file: prefer_const_constructors, non_constant_identifier_names, avoid_print, unused_import

import 'package:excel/func_excel.dart';
import 'package:flutter/material.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key});

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {

  List<int> lista_ids_tasks = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  List<String> nomes_funcs_excel = ['teste1', 'teste2', 'teste3', 'teste4', 'teste5', 'teste6', 'teste7', 'teste8', 'teste9', 'teste10'];
  List<String> vmp_excel = ['0.11', '0.22', '0.33', '0.44', '0.55', '0.66', '0.88', '0.99', '1.11', '2.11'];
  List<String> percent_erro_excel = ['0.28', '0.32', '0.45', '0.28', '0.44', '0.95', '1.5', '4.28', '0.53', '2.10'];

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Colors.blueGrey,
        title: Padding(
          padding: const EdgeInsets.only(left: 15),
          child: Text(
            "Flutter Excel",
            style: TextStyle(
              fontSize: 20,
              color: Colors.white,
              fontWeight: FontWeight.w500,
            ),
          ),
        ),
      ),
      body: Center(
        child: Column(
          children: [
            SizedBox(height: 500),
            GestureDetector(
              onTap: () {
                createExcel(lista_ids_tasks, nomes_funcs_excel, vmp_excel, percent_erro_excel);
              },
              child: Container(
                width: 200,
                height: 60,
                decoration: BoxDecoration(
                  borderRadius: BorderRadius.circular(12),
                  color: Colors.blue
                ),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceAround,
                  children: [
                    Center(
                      child: Text(
                        "Download Excel",
                        style: TextStyle(
                          color: Colors.white,
                          fontSize: 20,
                          fontWeight: FontWeight.w500,
                        ),
                      ),
                    ),
                    Icon(
                      Icons.download,
                      color: Colors.white,
                    )
                  ],
                ),
              ),
            )
          ],
        ),
      ),
    );
  }
}
