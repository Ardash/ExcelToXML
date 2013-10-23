using Kovalenko.Classes;
using System;

namespace TestExcelToXML {
  class Program {
    const String cmdLineError = "Не заданы параметры\n"+
                         "{0} <имя файла Excel> <имя XML файла для сохранения>";
    const String sPressAnyKey = "Нажмите любую клавишу...";
    const String processing = "Обрабатывается {0}";
    const String sOK = "Обработка завершена";

    static void pressAnyKey() {
      print(sPressAnyKey);
      Console.ReadKey();
    }

    static void print(String s) { 
      Console.WriteLine(s); 
    }

    static void Main(string[] args) {
      // Имя файла задано?
      if (args.Length < 2 || args[0].Length <= 0 || args[1].Length <= 0) {
        var exe = AppDomain.CurrentDomain.FriendlyName;
        print(String.Format(cmdLineError, exe));
        pressAnyKey();
        return;
      }

      // Обработка файла
      String ExcelName = args[0];
      String XMLName = args[1];
      print(String.Format(processing, ExcelName));
      String err = "";
      if (!libExcelToXML.ExcelToXML(ExcelName, XMLName, ref err)) print(err);
      else print(sOK);

      pressAnyKey();
    }
  }
}
