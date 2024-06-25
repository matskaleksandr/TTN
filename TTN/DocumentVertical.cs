using OfficeOpenXml;
using OfficeOpenXml.Style;
using Patagames.Pdf.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using static TTN.Table;

namespace TTN
{
    internal class DocumentVertical
    {
        //данные
        public string GruzOtpr = "";                                  //УНП грузоотправитель
        public string GruzPoluch = "";                                //УНП грузополучатель

        public string Date = "";                                      //Дата

        public string Avto = "";                                      //Автомобиль
        public string Pricep = "";                                    //Прицеп

        public string KPutList = "";                                  //К путевому листу номер

        public string Voditel = "";                                   //Водитель

        public string ZakazchikPerevozki = "";                        //Заказчик перевозки

        public string GruzOtprName = "";                              //Грузоотправитель название

        public string GruzPoluchName = "";                            //Грузоотправитель название

        public string OsnOtpusk = "";                                 //Основание отпуска

        public string PunktPogruzk = "";                              //Пункт погрузки
        public string PunktRazgruzki = "";                            //Пункт разгрузки

        public string Pereadresovka = "";                             //Переадресовка

        //Разделы
        public bool ShapkaYNPDate = false;                             //Шапка (УНП + дата)                 // 17

        public bool AutoInfo = false;                                 //Информация об автомобиле           // 10

        public bool GruzootpravitIOtpusk = false;                     //Грузоотправитель и отпуск          // 6

        public bool PunktPogruzkiPereeadresovki = false;              //Пункт погрузки и переадресовка     // 8        

        public bool StoimostIStoroni = false;                         //Стоимость и стороны                // 23

        public bool TovarnRazdel = false;                                                                  // 

        public bool PogruzRazgruz = false;                                                                 // 

        public bool ProchSved = false;                                                                     // 

        //данные
        public string VsegoSummNDS = "";                              //Всего сумма НДС
        public string VsegoSummNDSKop = "";

        public string VsegoStoimSNDS = "";                            //Всего стоимость с НДС
        public string VsegoStoimSNDSKop = "";

        public string VsegoMassGruz = "";                             //Всего масса груза

        public string OtpuskRazresh = "";                             //Отпуск разрешил

        public string SdalGruzootpav = "";                            //Сдал грузоотправитель

        public string NoPlomb = "";                                   //Номер пломбы

        public string VsegoKolGruzMest = "";                          //Всего количество грузовых мест

        public string TovarKPerevozkePrin = "";                       //Товар к перевозке принял

        public string PoDover = "";                                   //По доверенности

        public string Vidannoi = "";                                  //Выданной

        public string PrinGruzopoluch = "";                           //Принял грузополучатель

        public string NoPlomb2 = "";                                  //Номер промбы

        //таблицы
        public List<DataRazdel1> table1 = new List<DataRazdel1>();           //Данные товарного раздела

        public int globalRow = 1;

        public void CopyTable(int startRowForSecondTable, string filePath1 = "", string filePath2 = "", string outputPath = @"ExcelVertical\file1.xlsx")
        {
            using (var package1 = new ExcelPackage(new FileInfo(filePath1)))
            using (var package2 = new ExcelPackage(new FileInfo(filePath2)))
            using (var outputPackage = new ExcelPackage())
            {
                // Получаем первый лист из каждого файла
                var worksheet1 = package1.Workbook.Worksheets[0];
                var worksheet2 = package2.Workbook.Worksheets[0];

                // Создаем новый лист в выходном файле
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("MergedSheet");

                // Копируем объединение ячеек перед копированием содержимого и стилей
                CopyMergedCells(worksheet1, outputWorksheet, 0);
                CopyMergedCells(worksheet2, outputWorksheet, startRowForSecondTable - 1);

                // Копируем содержимое первого листа
                int rowCount1 = worksheet1.Dimension?.End.Row ?? 0;
                int colCount1 = worksheet1.Dimension?.End.Column ?? 0;
                for (int row = 1; row <= rowCount1; row++)
                {
                    for (int col = 1; col <= colCount1; col++)
                    {
                        var sourceCell = worksheet1.Cells[row, col];
                        var targetCell = outputWorksheet.Cells[row, col];

                        // Копируем значение
                        targetCell.Value = sourceCell.Value;

                        // Копируем стиль
                        CopyCellStyle(sourceCell.Style, targetCell.Style);
                    }
                }

                // Копируем содержимое второго листа, начиная с указанной строки
                int rowCount2 = worksheet2.Dimension?.End.Row ?? 0;
                int colCount2 = worksheet2.Dimension?.End.Column ?? 0;
                for (int row = 1; row <= rowCount2; row++)
                {
                    for (int col = 1; col <= colCount2; col++)
                    {
                        var sourceCell = worksheet2.Cells[row, col];
                        var targetCell = outputWorksheet.Cells[row + startRowForSecondTable - 1, col];

                        // Копируем значение
                        targetCell.Value = sourceCell.Value;

                        // Копируем стиль
                        CopyCellStyle(sourceCell.Style, targetCell.Style);
                    }
                }

                // Копируем размеры строк и столбцов для первого листа
                CopyRowHeights(worksheet1, outputWorksheet, rowCount1);
                CopyColumnWidths(worksheet1, outputWorksheet, colCount1);

                // Копируем размеры строк и столбцов для второго листа
                CopyRowHeights(worksheet2, outputWorksheet, rowCount2, startRowForSecondTable - 1);
                CopyColumnWidths(worksheet2, outputWorksheet, colCount2);

                // Сохраняем выходной файл
                outputPackage.SaveAs(new FileInfo(outputPath));
            }

        }
        public void ConvertToExcel(MainWindow main)
        {
            globalRow = 1;
            string filePath1 = "";
            string filePath2 = "";
            string outputPath = @"ExcelVertical\file1.xlsx";
            if (System.IO.File.Exists(@"ExcelVertical\file1.xlsx"))
            {
                System.IO.File.Delete(@"ExcelVertical\file1.xlsx");
            }
            System.IO.File.Copy(@"ExcelVertical\file.xlsx", @"ExcelVertical\file1.xlsx");

            bool isCheckedValue = main.cb1.IsChecked.HasValue && main.cb1.IsChecked.Value;
            ShapkaYNPDate = isCheckedValue;
            bool isCheckedValue2 = main.cb2.IsChecked.HasValue && main.cb2.IsChecked.Value;
            AutoInfo = isCheckedValue2;
            bool isCheckedValue3 = main.cb3.IsChecked.HasValue && main.cb3.IsChecked.Value;
            GruzootpravitIOtpusk = isCheckedValue3;
            bool isCheckedValue4 = main.cb4.IsChecked.HasValue && main.cb4.IsChecked.Value;
            PunktPogruzkiPereeadresovki = isCheckedValue4;
            bool isCheckedValue5 = main.cb6.IsChecked.HasValue && main.cb6.IsChecked.Value;
            StoimostIStoroni = isCheckedValue5;
            bool isCheckedValue6 = main.cb5.IsChecked.HasValue && main.cb5.IsChecked.Value;
            TovarnRazdel = isCheckedValue6;



            if (ShapkaYNPDate == true)
            {                
                filePath2 = @"ExcelVertical\Шапка.xlsx";
                filePath1 = @"ExcelVertical\file1.xlsx";
                CopyTable(globalRow, filePath1, filePath2, outputPath);
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells[globalRow + 3, 38].Value = GruzOtpr;
                    worksheet.Cells[globalRow + 3, 51].Value = GruzPoluch;
                    worksheet.Cells[globalRow + 15, 10].Value = Date;
                    package.Save();
                }
                globalRow += 17;
            }
            if (AutoInfo == true)
            {                
                filePath2 = @"ExcelVertical\Автомобиль.xlsx";
                filePath1 = @"ExcelVertical\file1.xlsx";
                CopyTable(globalRow, filePath1, filePath2, outputPath);
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells[globalRow, 14].Value = Avto;
                    worksheet.Cells[globalRow, 88].Value = Pricep;
                    worksheet.Cells[globalRow + 2, 21].Value = KPutList;
                    worksheet.Cells[globalRow + 4, 11].Value = Voditel;
                    worksheet.Cells[globalRow + 7, 45].Value = ZakazchikPerevozki;
                    package.Save();
                }
                globalRow += 10;
            }
            if (GruzootpravitIOtpusk == true)
            {                
                filePath2 = @"ExcelVertical\Грузоотправитель.xlsx";
                filePath1 = @"ExcelVertical\file1.xlsx";
                CopyTable(globalRow, filePath1, filePath2, outputPath);
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells[globalRow, 18].Value = GruzOtprName;
                    worksheet.Cells[globalRow + 2, 18].Value = GruzPoluchName;
                    worksheet.Cells[globalRow + 4, 17].Value = OsnOtpusk;
                    package.Save();
                }
                globalRow += 6;
            }
            if (PunktPogruzkiPereeadresovki == true)
            {
                filePath2 = @"ExcelVertical\Пункты.xlsx";
                filePath1 = @"ExcelVertical\file1.xlsx";
                CopyTable(globalRow, filePath1, filePath2, outputPath);
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells[globalRow, 17].Value = PunktPogruzk;
                    worksheet.Cells[globalRow, 87].Value = PunktRazgruzki;
                    worksheet.Cells[globalRow + 2, 16].Value = Pereadresovka;
                    package.Save();
                }
                globalRow += 8;
            }
            if(TovarnRazdel == true)
            {
                filePath2 = @"ExcelVertical\Товарная1.xlsx";
                filePath1 = @"ExcelVertical\file1.xlsx";
                CopyTable(globalRow, filePath1, filePath2, outputPath);
                globalRow += 3;
                for(int i = 0; i < table1.Count - 1; i++)
                {
                    filePath2 = @"ExcelVertical\Товарная2.xlsx";
                    filePath1 = @"ExcelVertical\file1.xlsx";
                    CopyTable(globalRow, filePath1, filePath2, outputPath);
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(outputPath)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        worksheet.Cells[globalRow, 1].Value = table1[i].НаименованиеТовара;
                        worksheet.Cells[globalRow, 32].Value = table1[i].ЕдиницаИзмерения;
                        worksheet.Cells[globalRow, 41].Value = table1[i].Количество;
                        worksheet.Cells[globalRow, 50].Value = table1[i].Цена;
                        worksheet.Cells[globalRow, 63].Value = table1[i].Стоимость;
                        worksheet.Cells[globalRow, 75].Value = table1[i].СтавкаНДС;
                        worksheet.Cells[globalRow, 82].Value = table1[i].СуммаНДС;
                        worksheet.Cells[globalRow, 93].Value = table1[i].СтоимостьСНДС;
                        worksheet.Cells[globalRow, 105].Value = "";
                        worksheet.Cells[globalRow, 114].Value = "";
                        worksheet.Cells[globalRow, 124].Value = table1[i].Примечание;
                        package.Save();
                    }
                    globalRow += 1;
                }
                filePath2 = @"ExcelVertical\Товарная3.xlsx";
                filePath1 = @"ExcelVertical\file1.xlsx";
                CopyTable(globalRow, filePath1, filePath2, outputPath);
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells[globalRow, 41].Value = table1[table1.Count - 1].Количество;
                    worksheet.Cells[globalRow, 63].Value = table1[table1.Count - 1].Стоимость;
                    worksheet.Cells[globalRow, 82].Value = table1[table1.Count - 1].СуммаНДС;
                    worksheet.Cells[globalRow, 93].Value = table1[table1.Count - 1].СтоимостьСНДС;
                    package.Save();
                }
                globalRow += 1;
            }
            if (StoimostIStoroni == true)
            {
                filePath2 = @"ExcelVertical\СтоимостьИСтороны.xlsx";
                filePath1 = @"ExcelVertical\file1.xlsx";
                CopyTable(globalRow, filePath1, filePath2, outputPath);
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputPath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells[globalRow, 19].Value = VsegoSummNDS;
                    worksheet.Cells[globalRow, 108].Value = VsegoSummNDSKop;

                    worksheet.Cells[globalRow + 3, 23].Value = VsegoStoimSNDS;
                    worksheet.Cells[globalRow + 3, 108].Value = VsegoStoimSNDSKop;

                    worksheet.Cells[globalRow + 5, 15].Value = VsegoMassGruz;
                    worksheet.Cells[globalRow + 5, 87].Value = VsegoKolGruzMest;

                    worksheet.Cells[globalRow + 7, 15].Value = OtpuskRazresh;
                    worksheet.Cells[globalRow + 11, 19].Value = SdalGruzootpav;
                    worksheet.Cells[globalRow + 16, 10].Value = NoPlomb;

                    worksheet.Cells[globalRow + 7, 83].Value = TovarKPerevozkePrin;
                    worksheet.Cells[globalRow + 11, 76].Value = PoDover;
                    worksheet.Cells[globalRow + 13, 71].Value = Vidannoi;
                    worksheet.Cells[globalRow + 15, 82].Value = PrinGruzopoluch;
                    worksheet.Cells[globalRow + 19, 74].Value = NoPlomb2;

                    package.Save();
                }
                globalRow += 23;
            }

            System.IO.Directory.CreateDirectory(@"Exits");
            if (System.IO.File.Exists(@"Exits\file.xlsx"))
            {
                System.IO.File.Delete(@"Exits\file.xlsx");
            }
            System.IO.File.Copy(@"ExcelVertical\file1.xlsx", @"Exits\file.xlsx");
        }

        static void CopyCellStyle(ExcelStyle sourceStyle, ExcelStyle targetStyle)
        {
            targetStyle.Font.Bold = sourceStyle.Font.Bold;
            targetStyle.Font.Italic = sourceStyle.Font.Italic;
            targetStyle.Font.Size = sourceStyle.Font.Size;
            targetStyle.HorizontalAlignment = sourceStyle.HorizontalAlignment;
            targetStyle.VerticalAlignment = sourceStyle.VerticalAlignment;
            targetStyle.Font.Name = "Times New Roman";
            targetStyle.WrapText = sourceStyle.WrapText;

            if (sourceStyle.Fill.PatternType == ExcelFillStyle.Solid)
            {
                targetStyle.Fill.PatternType = ExcelFillStyle.Solid;
                if (sourceStyle.Fill.BackgroundColor.Rgb != null)
                {
                    targetStyle.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                }
            }

            // Копируем границы
            CopyBorderStyle(sourceStyle.Border.Top, targetStyle.Border.Top);
            CopyBorderStyle(sourceStyle.Border.Bottom, targetStyle.Border.Bottom);
            CopyBorderStyle(sourceStyle.Border.Left, targetStyle.Border.Left);
            CopyBorderStyle(sourceStyle.Border.Right, targetStyle.Border.Right);
        }

        static void CopyBorderStyle(ExcelBorderItem sourceBorder, ExcelBorderItem targetBorder)
        {
            targetBorder.Style = sourceBorder.Style;
        }

        static void CopyRowHeights(ExcelWorksheet sourceWorksheet, ExcelWorksheet targetWorksheet, int rowCount, int startRowOffset = 0)
        {
            for (int row = 1; row <= rowCount; row++)
            {
                if (sourceWorksheet.Row(row).Height > 0)
                {
                    targetWorksheet.Row(row + startRowOffset).Height = sourceWorksheet.Row(row).Height;
                }
            }
        }

        static void CopyColumnWidths(ExcelWorksheet sourceWorksheet, ExcelWorksheet targetWorksheet, int colCount)
        {
            for (int col = 1; col <= colCount; col++)
            {
                if (sourceWorksheet.Column(col).Width > 0)
                {
                    targetWorksheet.Column(col).Width = sourceWorksheet.Column(col).Width;
                }
            }
        }

        static void CopyMergedCells(ExcelWorksheet sourceWorksheet, ExcelWorksheet targetWorksheet, int rowOffset)
        {
            foreach (var mergedCell in sourceWorksheet.MergedCells)
            {
                var cellAddresses = new ExcelAddress(mergedCell);
                var startRow = cellAddresses.Start.Row + rowOffset;
                var endRow = cellAddresses.End.Row + rowOffset;
                var startCol = cellAddresses.Start.Column;
                var endCol = cellAddresses.End.Column;

                targetWorksheet.Cells[startRow, startCol, endRow, endCol].Merge = true;
            }
        }

    }
}
