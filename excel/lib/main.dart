// ignore_for_file: prefer_const_constructors
// ignore_for_file: non_constant_identifier_names
// ignore_for_file: avoid_print
// ignore_for_file: unused_import
// ignore_for_file:
// ignore_for_file:
// ignore_for_file:
// ignore_for_file:
// ignore_for_file:
// ignore_for_file:
// ignore_for_file:

//imports
import 'package:excel/listed_sheet.dart';
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

  Color corCt_1 = Color.fromARGB(255, 199, 192, 192);
  Color corCt_2 = Color.fromARGB(255, 199, 192, 192);
  Color corCt_3 = Color.fromARGB(255, 199, 192, 192);
  Color corCt_4 = Color.fromARGB(255, 199, 192, 192);
  Color corCt_5 = Color.fromARGB(255, 199, 192, 192);
  Color corCt_6 = Color.fromARGB(255, 199, 192, 192);
  Color corCt_7 = Color.fromARGB(255, 199, 192, 192);
  Color corCt_8 = Color.fromARGB(255, 199, 192, 192);

  @override
  Widget build(BuildContext context) {

    // Screen dimensions
    double screenWidth = MediaQuery.of(context).size.width;
    double screenHeight = MediaQuery.of(context).size.height;

    return Scaffold(
      appBar: AppBar(
        backgroundColor: Colors.blueGrey,
        title: Padding(
          padding: const EdgeInsets.only(left: 15),
          child: Text(
            "Flutter OfficeChart RoadMap",
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
            Container(
              height: screenHeight / 2.14,
              color: Colors.transparent,
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  //ct1 sup
                  Padding(
                    padding: const EdgeInsets.only(left: 40),
                    child: MouseRegion(
                      onHover: (_) {
                        setState(() {
                          corCt_1 = Color.fromARGB(255, 139, 135, 135);
                        });
                      },
                      onExit: (_) {
                        setState(() {
                          corCt_1 = Color.fromARGB(255, 199, 192, 192);
                        });
                      },
                      child: GestureDetector(
                        onTap: () {
                          setState(() {
                            listed_sheet();
                          });
                        },
                        child: Container(
                          width: screenWidth / 4.5,
                          height: screenHeight / 2.4,
                          decoration: BoxDecoration(
                            borderRadius: BorderRadius.circular(16),
                            color: corCt_1,
                          ),
                          child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Listed Sheet",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                        ),
                      ),
                    ),
                  ),
                  //ct2 sup
                  MouseRegion(
                    onHover: (_) {
                      setState(() {
                        corCt_2 = Color.fromARGB(255, 139, 135, 135);
                      });
                    },
                    onExit: (_) {
                      setState(() {
                        corCt_2 = Color.fromARGB(255, 199, 192, 192);
                      });
                    },
                    child: GestureDetector(
                      onTap: () {
                        setState(() {
                          //
                        });
                      },
                      child: Container(
                        width: screenWidth / 4.5,
                        height: screenHeight / 2.4,
                        decoration: BoxDecoration(
                          borderRadius: BorderRadius.circular(16),
                          color: corCt_2,
                        ),
                        child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Pie Chart",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                      ),
                    ),
                  ),
                  //ct3 sup
                  MouseRegion(
                    onHover: (_) {
                      setState(() {
                        corCt_3 = Color.fromARGB(255, 139, 135, 135);
                      });
                    },
                    onExit: (_) {
                      setState(() {
                        corCt_3 = Color.fromARGB(255, 199, 192, 192);
                      });
                    },
                    child: GestureDetector(
                      onTap: () {
                        setState(() {
                          //
                        });
                      },
                      child: Container(
                        width: screenWidth / 4.5,
                        height: screenHeight / 2.4,
                        decoration: BoxDecoration(
                          borderRadius: BorderRadius.circular(16),
                          color: corCt_3,
                        ),
                        child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Bar Chart",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                      ),
                    ),
                  ),
                  //ct4 sup
                  Padding(
                    padding: const EdgeInsets.only(right: 40),
                    child: MouseRegion(
                      onHover: (_) {
                        setState(() {
                          corCt_4 = Color.fromARGB(255, 139, 135, 135);
                        });
                      },
                      onExit: (_) {
                        setState(() {
                          corCt_4 = Color.fromARGB(255, 199, 192, 192);
                        });
                      },
                      child: GestureDetector(
                        onTap: () {
                          setState(() {
                            //
                          });
                        },
                        child: Container(
                          width: screenWidth / 4.5,
                          height: screenHeight / 2.4,
                          decoration: BoxDecoration(
                            borderRadius: BorderRadius.circular(16),
                            color: corCt_4,
                          ),
                          child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Dual Bar Chart",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                        ),
                      ),
                    ),
                  ),
                ],
              ),
            ),
            Container(
              height: screenHeight / 2.14,
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  //ct1 inf
                  Padding(
                    padding: const EdgeInsets.only(left: 40),
                    child: MouseRegion(
                      onHover: (_) {
                        setState(() {
                          corCt_5 = Color.fromARGB(255, 139, 135, 135);
                        });
                      },
                      onExit: (_) {
                        setState(() {
                          corCt_5 = Color.fromARGB(255, 199, 192, 192);
                        });
                      },
                      child: GestureDetector(
                        onTap: () {
                          setState(() {
                            //
                          });
                        },
                        child: Container(
                          width: screenWidth / 4.5,
                          height: screenHeight / 2.4,
                          decoration: BoxDecoration(
                            borderRadius: BorderRadius.circular(16),
                            color: corCt_5,
                          ),
                          child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Line Chart",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                        ),
                      ),
                    ),
                  ),
                  //ct2 inf
                  MouseRegion(
                    onHover: (_) {
                      setState(() {
                        corCt_6 = Color.fromARGB(255, 139, 135, 135);
                      });
                    },
                    onExit: (_) {
                      setState(() {
                        corCt_6 = Color.fromARGB(255, 199, 192, 192);
                      });
                    },
                    child: GestureDetector(
                      onTap: () {
                        setState(() {
                          //
                        });
                      },
                      child: Container(
                        width: screenWidth / 4.5,
                        height: screenHeight / 2.4,
                        decoration: BoxDecoration(
                          borderRadius: BorderRadius.circular(16),
                          color: corCt_6,
                        ),
                        child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Stacked Chart",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                      ),
                    ),
                  ),
                  //ct3 inf
                  MouseRegion(
                    onHover: (_) {
                      setState(() {
                        corCt_7 = Color.fromARGB(255, 139, 135, 135);
                      });
                    },
                    onExit: (_) {
                      setState(() {
                        corCt_7 = Color.fromARGB(255, 199, 192, 192);
                      });
                    },
                    child: GestureDetector(
                      onTap: () {
                        setState(() {
                          //
                        });
                      },
                      child: Container(
                        width: screenWidth / 4.5,
                        height: screenHeight / 2.4,
                        decoration: BoxDecoration(
                          borderRadius: BorderRadius.circular(16),
                          color: corCt_7,
                        ),
                        child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Stacked Bar Chart",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                      ),
                    ),
                  ),
                  //ct4
                  Padding(
                    padding: const EdgeInsets.only(right: 40),
                    child: MouseRegion(
                      onHover: (_) {
                        setState(() {
                          corCt_8 = Color.fromARGB(255, 139, 135, 135);
                        });
                      },
                      onExit: (_) {
                        setState(() {
                          corCt_8 = Color.fromARGB(255, 199, 192, 192);
                        });
                      },
                      child: GestureDetector(
                        onTap: () {
                          setState(() {
                            //
                          });
                        },
                        child: Container(
                          width: screenWidth / 4.5,
                          height: screenHeight / 2.4,
                          decoration: BoxDecoration(
                            borderRadius: BorderRadius.circular(16),
                            color: corCt_8,
                          ),
                          child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              SizedBox(height: 20),
                              Center(
                                child: Text(
                                  "Stacked Line Chart",
                                  style: TextStyle(
                                      fontSize: 32,
                                      fontWeight: FontWeight.w500,
                                      color: Colors.black),
                                ),
                              ),
                              SizedBox(height: 20),
                              Icon(
                                Icons.download,
                                color: Colors.black,
                                size: 40,
                              )
                            ],
                          ),
                        ),
                      ),
                    ),
                  ),
                ],
              ),
            )
          ],
        ),
      ),
    );
  }
}
