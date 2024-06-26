using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Tesseract;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using DuoVia.FuzzyStrings;


using System.Text.RegularExpressions;
using System.Windows.Markup;
using Aspose.Pdf;
using System.Windows.Media.TextFormatting;
using System.Reflection.Emit;
using Aspose.Pdf.Vector;
using static TTN.Table;

namespace TTN
{
    public partial class MainWindow : Window
    {
        public string pathfile;
        public string excelFilePath = "результат.xlsx";
        public string textt = "";
        DocumentVertical documentV;
        public bool vertical;
        public ImageBrush imgDelete;
        public List<string> typesData = new List<string>() {
            "УНП.Грузоотправитель",
            "УНП.Грузополучатель",
            "УНП.ЗаказчикАвтомобильнойПеревозки",
            "Шапка.Дата",
            "Шапка.Грузоотправитель",
            "Шапка.Грузополучатель",
            "Шапка.ОснованиеОтпуска",
            "ВсегоСуммаНДС",
            "ВсегоCтоимостьCНДС",
            "ОтпускРазрешил",
            "СдалГрузоотправитель",
            "ТоварКПеревозкеПринял",
            "ПоДоверенности(#)",
            "ДоверенностьВыдана",
            "1ТОВАРНЫЙРАЗДЕЛ",
        };
        public List<Grid> grid = new List<Grid>();
        Table tb = null;
        List<DataRazdel1> items = new List<DataRazdel1>();




        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            excelFilePath = System.IO.Path.Combine(exePath, excelFilePath);
            comboBoxDataTypes.ItemsSource = typesData;
            imgDelete = brush;
        }
        private void AddData(int type, string data)
        {
            Grid clonedGrid = CreateDuplicatedGrid(type, data);
            MainStackPanel.Children.Add(clonedGrid);
            grid.Add(clonedGrid);
            clonedGrid.Visibility = Visibility.Visible;
            if (grid.Count != 0 && (documentV != null || documentV != null))
            {
                buttonExcel.IsEnabled = true;
                buttonXML.IsEnabled = true;
            }
        }


        public void TesseractStart(object sender, RoutedEventArgs e)
        {
            Filter filter = new Filter(ochist);
            string outputDirectory = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ConvertedImages");
            filter.FilterWhite(pathfile, outputDirectory);

            using (var engine = new TesseractEngine(@"D:\Programm\editor\Tesseract-OCR\tessdata", "eng+rus", EngineMode.Default))
            {
                using (var image = Pix.LoadFromFile(Path.Combine(outputDirectory, $"doc1.png")))
                {
                    using (var page = engine.Process(image, PageSegMode.Auto))
                    {
                        using (var package = new ExcelPackage())
                        {
                            var worksheet = package.Workbook.Worksheets.Add("OCR Results");
                            using (var iterator = page.GetIterator())
                            {
                                iterator.Begin();
                                int row = 1;
                                double MaxX = 0;
                                int nRow = 1;
                                int nColumn = 1;
                                Bitmap originalImage = new Bitmap(Path.Combine(outputDirectory, $"doc1.png"));
                                Bitmap copiedImage = new Bitmap(originalImage.Width, originalImage.Height);

                                int width = originalImage.Width;
                                int height = originalImage.Height;

                                if (width > height)
                                {
                                    vertical = false;
                                }
                                else if (height > width)
                                {
                                    vertical = true;
                                    documentV = new DocumentVertical();
                                }

                                do
                                {
                                    string currentWord = iterator.GetText(PageIteratorLevel.Word);
                                    iterator.TryGetBoundingBox(PageIteratorLevel.Word, out Tesseract.Rect bounds);

                                    if (bounds.X1 < MaxX)
                                    {
                                        MaxX = bounds.X1;
                                        nRow++;
                                    }
                                    else
                                    {
                                        MaxX = bounds.X1;
                                    }



                                    worksheet.Cells[nRow, nColumn].Value += " " + currentWord;
                                    bool viv = false;

                                    if (currentWord != null && currentWord.Length != 0 && currentWord != "" && currentWord != " " && viv == true)
                                    {
                                        for (int x = 0; x < originalImage.Width; x++)
                                        {
                                            for (int y = 0; y < originalImage.Height; y++)
                                            {
                                                if (x >= bounds.X1 && x <= bounds.X2 && (y == bounds.Y1 || y == bounds.Y2))
                                                {
                                                    System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                                                    if (pixelColor != System.Drawing.Color.Green)
                                                    {
                                                        copiedImage.SetPixel(x, y, System.Drawing.Color.Red);
                                                        //MessageBox.Show("");
                                                    }
                                                }
                                                if (y >= bounds.Y1 && y <= bounds.Y2 && (x == bounds.X1 || x == bounds.X2))
                                                {
                                                    System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                                                    if (pixelColor != System.Drawing.Color.Green)
                                                    {
                                                        copiedImage.SetPixel(x, y, System.Drawing.Color.Red);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (currentWord == null || currentWord.Length == 0 || currentWord == "" || currentWord == " " && viv == true)
                                    {
                                        //MessageBox.Show($"X1{bounds.X1}, Y1{bounds.Y1}, X2{bounds.X2}, Y2{bounds.Y2}");
                                        for (int x = 0; x < originalImage.Width; x++)
                                        {
                                            for (int y = 0; y < originalImage.Height; y++)
                                            {
                                                if (x >= bounds.X1 && x <= bounds.X2 && (y == bounds.Y1 || y == bounds.Y2))
                                                {
                                                    System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                                                    if (pixelColor != System.Drawing.Color.Green)
                                                    {
                                                        copiedImage.SetPixel(x, y, System.Drawing.Color.Blue);
                                                    }
                                                }
                                                if (y >= bounds.Y1 && y <= bounds.Y2 && (x == bounds.X1 || x == bounds.X2))
                                                {
                                                    System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                                                    if (pixelColor != System.Drawing.Color.Green)
                                                    {
                                                        copiedImage.SetPixel(x, y, System.Drawing.Color.Blue);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    copiedImage.SetPixel(bounds.X1, bounds.Y1, System.Drawing.Color.Green);
                                    copiedImage.SetPixel(bounds.X2, bounds.Y2, System.Drawing.Color.Green);
                                    copiedImage.SetPixel(bounds.X1, bounds.Y2, System.Drawing.Color.Green);
                                    copiedImage.SetPixel(bounds.X2, bounds.Y1, System.Drawing.Color.Green);
                                    row++;
                                } while (iterator.Next(PageIteratorLevel.Word));
                                copiedImage.Save(Path.Combine(outputDirectory, $"doc2.png"));
                                originalImage.Dispose();
                                copiedImage.Dispose();
                            }
                            package.SaveAs(new FileInfo(excelFilePath));
                        }
                    }
                }
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(excelFilePath);
        }
        //устаревший но работающий код/

        //хороший код\
        public void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Изображения (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp|Все файлы (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string selectedFilePath = openFileDialog.FileName;
                BitmapImage bitmap = new BitmapImage();
                documentV = null;
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(selectedFilePath, UriKind.RelativeOrAbsolute);
                bitmap.EndInit();
                imgBox.Source = bitmap;
                pathfile = selectedFilePath;
                viewbox1.Visibility = Visibility.Hidden;
                //изменение состояния кнопок
                buttonScan.IsEnabled = true;
                menuButtonScan.IsEnabled = true;
                buttonZoom.Visibility = Visibility.Visible;
                buttonZoom2.Visibility = Visibility.Visible;
                if (bitmap.Width > bitmap.Height)
                {

                }
                else
                {
                    documentV = new DocumentVertical();

                }
            }
        }
        List<Tuple<int, int, int, int>> lineCoordinates = new List<Tuple<int, int, int, int>>();
        List<Tables> tables = new List<Tables>();
        private void DetermineEdgeType(List<Tuple<int, int, int, int>> lines)
        {
            List<Tuple<int, int, int, int>> horizontalLines = new List<Tuple<int, int, int, int>>();
            List<Tuple<int, int, int, int>> verticalLines = new List<Tuple<int, int, int, int>>();

            foreach (var line in lines)
            {
                int B1 = line.Item3 - line.Item1; //разница по X
                int B2 = line.Item4 - line.Item2; //разница по Y
                if (B1 > B2)
                {
                    int B3 = (line.Item2 + line.Item4) / 2;
                    horizontalLines.Add(Tuple.Create(line.Item1, B3, line.Item3, B3));
                }
                else
                {
                    int B3 = (line.Item1 + line.Item3) / 2;
                    if (line.Item3 - line.Item1 <= 50)
                    {
                        verticalLines.Add(Tuple.Create(B3, line.Item2, B3, line.Item4));
                    }
                }
            }
            foreach (var lineH in horizontalLines)
            {
                foreach (var lineV in verticalLines)
                {
                    if (Math.Abs(lineH.Item1 - lineV.Item1) < 100)
                    {
                        if (Math.Abs(lineH.Item2 - lineV.Item2) < 100)
                        {
                            Tables table = new Tables(); //новая таблица
                            table.KORDx.Add(lineH.Item1);
                            table.KORDy.Add(lineV.Item2);

                            foreach (var lineH2 in horizontalLines)
                            {
                                if ((lineH2.Item2 >= lineV.Item2 && lineH2.Item2 <= lineV.Item4) && lineH != lineH2)
                                {
                                    if (Math.Abs(lineH2.Item1 - lineV.Item1) < 50)
                                    {
                                        table.KORDy.Add(lineH2.Item2);
                                        table.KORDy.Sort();
                                    }
                                }
                            }
                            tables.Add(table);
                        }
                    }
                    else if ((lineV.Item1 - 20 >= lineH.Item1 && lineV.Item1 <= lineH.Item3 + 20) && tables.Count != 0) // XV находится между X12H
                    {
                        bool r = false;
                        if (Math.Abs(tables[tables.Count - 1].KORDy[0] - lineV.Item2) < 100)
                        {
                            tables[tables.Count - 1].KORDx.Add(lineV.Item1);
                            tables[tables.Count - 1].KORDx.Sort();
                            r = true;
                        }
                        if (Math.Abs(lineH.Item3 - lineV.Item1) < 50)
                        {
                            if (Math.Abs(tables[tables.Count - 1].KORDy[0] - lineV.Item2) < 200 && r == false)
                            {
                                tables[tables.Count - 1].KORDx.Add(lineV.Item1);
                                tables[tables.Count - 1].KORDx.Sort();
                            }
                        }
                    }
                }
            }
            foreach (var table in tables)
            {
                int z = 0;
                for (int i = 0; i < table.KORDx.Count; i++)
                {
                    if (table.KORDx[i] != z)
                    {
                        z = table.KORDx[i];
                    }
                    else
                    {
                        table.KORDx.Remove(z);
                    }
                }
                int u = 0;
                for (int i = 0; i < table.KORDy.Count; i++)
                {
                    if (table.KORDy[i] != u)
                    {
                        u = table.KORDy[i];
                    }
                    else
                    {
                        table.KORDy.Remove(u);
                    }
                }
            }
        }//обработка таблиц
        public async Task Scan()
        {
            if(grid.Count != 0)
            {
                for(int i = 0; i < grid.Count; i++)
                {
                    MainStackPanel.Children.Remove(grid[i]);
                }                
            }
            grid.Clear();
            if (documentV != null)
            {
                documentV = null;
                documentV = new DocumentVertical();
            }
            tables.Clear();
            tables = new List<Tables>();
            lineCoordinates.Clear();
            lineCoordinates = new List<Tuple<int, int, int, int>>();

            if (items.Count != 0)
            {
                items.Clear();
                items = new List<DataRazdel1>();
            }
            System.IO.Directory.CreateDirectory(@"ConvertedImages");
            string outputDirectory = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ConvertedImages");


            await Task.Run(() =>
            {
                
                bool isChecked = false;
                Dispatcher.Invoke(() => isChecked = checkFilter.IsChecked == true);
                Dispatcher.Invoke(() => progressBar.Value = 5);
                Dispatcher.Invoke(() => buttonScan.IsEnabled = false);

                Filter filter = new Filter(ochist);
                filter.FilterWhite(pathfile, outputDirectory);

                Dispatcher.Invoke(() => progressBar.Value = 15);

                List<string> rows = new List<string>();
                List<List<string>> listOfRows = new List<List<string>>();

                bool tablecheck = false;

                using (var engine = new TesseractEngine(@"D:\Programm\editor\Tesseract-OCR\tessdata", "eng+rus", EngineMode.Default))
                {
                    using (var image = Pix.LoadFromFile(Path.Combine(outputDirectory, $"doc1.png")))
                    {
                        using (var page = engine.Process(image, PageSegMode.Auto))
                        {
                            using (var package = new ExcelPackage())
                            {
                                using (var iterator = page.GetIterator())
                                {
                                    iterator.Begin();
                                    int row = 1;
                                    double MaxX = 0;
                                    double MaxY = 0;
                                    int nRow = 1;
                                    Bitmap originalImage = new Bitmap(Path.Combine(outputDirectory, $"doc1.png"));
                                    Bitmap copiedImage = new Bitmap(originalImage.Width, originalImage.Height);

                                    int width = originalImage.Width;
                                    int height = originalImage.Height;

                                    if (width > height)
                                    {
                                        vertical = false;
                                    }
                                    else if (height > width)
                                    {
                                        vertical = true;
                                        documentV = new DocumentVertical();
                                    }

                                    do
                                    {
                                        string currentWord = iterator.GetText(PageIteratorLevel.Word);
                                        iterator.TryGetBoundingBox(PageIteratorLevel.Word, out Tesseract.Rect bounds);
                                        string text = string.Join("", rows);
                                        if (text.IndexOf("ТОВАРНЫЙ", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("РАЗДЕЛ", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            tablecheck = true;
                                        }
                                        if (tablecheck == true)
                                        {
                                            ScanTable(currentWord, bounds);
                                        }
                                        if (currentWord == "|")
                                        {
                                            currentWord = null;
                                        }
                                        if (bounds.X1 < MaxX)
                                        {
                                            MaxX = bounds.X1;
                                            nRow++;
                                            rows = new List<string>();
                                        }
                                        else if (Math.Abs((float)((bounds.Y1 + bounds.Y2) / 2) - MaxY) > 200)
                                        {
                                            MaxY = (float)((bounds.Y1 + bounds.Y2) / 2);
                                            nRow++;
                                            rows = new List<string>();
                                        }
                                        else
                                        {
                                            MaxX = bounds.X1;
                                            MaxY = (float)((bounds.Y1 + bounds.Y2) / 2);
                                        }
                                        rows.Add(" " + currentWord);
                                        if (listOfRows.Count != nRow)
                                        {
                                            listOfRows.Add(rows);
                                        }
                                        else
                                        {
                                            listOfRows[nRow - 1] = rows;
                                        }
                                        if (1 > 1)
                                        {
                                            DebugTesseractZone(currentWord, originalImage, bounds, copiedImage);
                                        }
                                        row++;
                                    } while (iterator.Next(PageIteratorLevel.Word));
                                    copiedImage.Save(Path.Combine(outputDirectory, $"doc2.png"));
                                    originalImage.Dispose();
                                    copiedImage.Dispose();
                                }
                                //package.SaveAs(new FileInfo(excelFilePath));
                            }
                        }
                    }
                }
                Dispatcher.Invoke(() => progressBar.Value = 50);
                bool boolDateHead = false;
                bool boolGruzootpav = false;
                bool boolGruzopoluch = false;
                bool boolOsnovanOtpusk = false;
                bool boolVsegoSummaNDS = false;
                bool boolVsegoStoimostSNDS = false;
                bool boolOtpuskRazreshil = false;
                bool boolSdalGruzootpravit = false;
                bool boolTovarKDostavkePrin = false;
                bool boolPoDoverenn = false;
                bool boolDoverennVidana = false;
                bool boolTOVARNRAZDEL = false;

                DetermineEdgeType(lineCoordinates);
                //MessageBox.Show(tables.Count.ToString());
                Bitmap originalImage2 = new Bitmap(Path.Combine(outputDirectory, $"doc1.png"));
                Bitmap copiedImage2 = new Bitmap(originalImage2.Width, originalImage2.Height);

                if (tables.Count > 0)
                {
                    foreach (var table in tables)
                    {
                        for (int i = table.KORDx[0]; i < table.KORDx[table.KORDx.Count - 1]; i++)
                        {
                            foreach (var y in table.KORDy)
                            {
                                copiedImage2.SetPixel(i, y, System.Drawing.Color.Yellow);
                            }
                        }
                        for (int i = table.KORDy[0]; i < table.KORDy[table.KORDy.Count - 1]; i++)
                        {
                            foreach (var x in table.KORDx)
                            {
                                copiedImage2.SetPixel(x, i, System.Drawing.Color.Yellow);
                            }
                        }
                    }
                    foreach (var table in tables)
                    {
                        foreach (var x in table.KORDx)
                        {
                            foreach (var y in table.KORDy)
                            {
                                copiedImage2.SetPixel(x, y, System.Drawing.Color.White);
                            }
                        }
                    }
                }

                copiedImage2.Save(Path.Combine(outputDirectory, $"doc3.png"));
                originalImage2.Dispose();
                copiedImage2.Dispose();
                Dispatcher.Invoke(() => progressBar.Value = 65);
                for (int i = 0; i < listOfRows.Count; i++)
                {

                    for (int j = 0; j < listOfRows[i].Count; j++)
                    {
                        string currentWord = listOfRows[i][j];

                        if (currentWord.IndexOf("Грузоотправитель", StringComparison.OrdinalIgnoreCase) >= 0 && i < 8)
                        {
                            string data = null;
                            foreach (var row in listOfRows)
                            {
                                string line = string.Join("", row);
                                if (line.IndexOf("УНП", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    string[] lineData = line.Split();
                                    data = lineData[2];
                                    Dispatcher.Invoke(() => cb1.IsChecked = true);
                                    break;
                                }
                            }
                            if (documentV != null)
                            {
                                documentV.GruzOtpr = data;
                            }
                            Dispatcher.Invoke(() => AddData(0, data));
                        }
                        if (currentWord.IndexOf("Грузополучатель", StringComparison.OrdinalIgnoreCase) >= 0 && i < 8)
                        {
                            string data = null;
                            foreach (var row in listOfRows)
                            {
                                string line = string.Join("", row);
                                if (line.IndexOf("УНП", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    string[] lineData = line.Split();
                                    data = lineData[3];
                                    Dispatcher.Invoke(() => cb1.IsChecked = true);
                                    break;
                                }
                            }
                            if (documentV != null)
                            {
                                documentV.GruzPoluch = data;
                            }
                            Dispatcher.Invoke(() => AddData(1, data));
                        }
                        if (boolDateHead == false)
                        {
                            string pattern = @"\b\d{1,2}\s(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s\d{4}\b";
                            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
                            Match match = regex.Match(string.Join("", listOfRows[i]));
                            if (match.Success)
                            {
                                if (documentV != null)
                                {
                                    documentV.Date = match.Value;
                                }
                                Dispatcher.Invoke(() => AddData(3, match.Value));
                                Dispatcher.Invoke(() => cb1.IsChecked = true);
                                boolDateHead = true;
                            }
                        }
                        if (boolGruzootpav == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            List<string> textList = text.Split().ToList();
                            for (int k = 0; k < textList.Count; k++)
                            {
                                if (textList[k] == " " || textList[k] == "")
                                {
                                    textList.Remove(textList[k]);
                                }
                            }
                            if (currentWord.IndexOf("Грузоотправитель", StringComparison.OrdinalIgnoreCase) >= 0 && textList.Count > 4)
                            {
                                if (documentV != null)
                                {
                                    documentV.GruzOtprName = RemoveFirstWord(text, 1);
                                }
                                Dispatcher.Invoke(() => cb3.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(4, RemoveFirstWord(text, 1)));
                                boolGruzootpav = true;
                            }
                        }
                        if (boolGruzopoluch == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            List<string> textList = text.Split().ToList();
                            for (int k = 0; k < textList.Count; k++)
                            {
                                if (textList[k] == " " || textList[k] == "")
                                {
                                    textList.Remove(textList[k]);
                                }
                            }
                            if (currentWord.IndexOf("Грузополучатель", StringComparison.OrdinalIgnoreCase) >= 0 && textList.Count > 4)
                            {
                                if (documentV != null)
                                {
                                    documentV.GruzPoluchName = RemoveFirstWord(text, 2);
                                }
                                Dispatcher.Invoke(() => cb3.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(5, RemoveFirstWord(text)));
                                boolGruzopoluch = true;
                            }
                        }
                        if (boolOsnovanOtpusk == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            if (text.IndexOf("Основание", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("отпуска", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (documentV != null)
                                {
                                    documentV.OsnOtpusk = RemoveFirstWord(text, 2);
                                }
                                Dispatcher.Invoke(() => cb3.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(6, RemoveFirstWord(text, 2)));
                                boolOsnovanOtpusk = true;
                            }
                        }
                        if (boolVsegoSummaNDS == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            if (text.IndexOf("Всего", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("сумма", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("НДС", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (documentV != null)
                                {
                                    documentV.VsegoSummNDS = RemoveFirstWord(text, 3);
                                }
                                Dispatcher.Invoke(() => cb6.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(7, RemoveFirstWord(text, 3)));
                                boolVsegoSummaNDS = true;
                            }
                        }
                        if (boolVsegoStoimostSNDS == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            if (text.IndexOf("Всего", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("стоимость", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("с", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("НДС", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (documentV != null)
                                {
                                    documentV.VsegoStoimSNDS = RemoveFirstWord(text, 4);
                                }
                                Dispatcher.Invoke(() => cb6.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(8, RemoveFirstWord(text, 4)));
                                boolVsegoStoimostSNDS = true;
                            }
                        }
                        if (boolOtpuskRazreshil == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            if (text.IndexOf("Отпуск", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("разрешил", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (documentV != null)
                                {
                                    documentV.OtpuskRazresh = RemoveFirstWord(text, 2);
                                }
                                Dispatcher.Invoke(() => cb6.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(9, RemoveFirstWord(text, 2)));
                                boolOtpuskRazreshil = true;
                            }
                        }
                        if (boolSdalGruzootpravit == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            if (text.IndexOf("Сдал", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("Грузоотправитель", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (documentV != null)
                                {
                                    documentV.SdalGruzootpav = RemoveFirstWord(text, 2);
                                }
                                Dispatcher.Invoke(() => cb6.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(10, RemoveFirstWord(text, 2)));
                                boolSdalGruzootpravit = true;
                            }
                        }
                        if (boolTovarKDostavkePrin == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            if (text.IndexOf("Товар", StringComparison.OrdinalIgnoreCase) >= 0
                                && text.IndexOf("к", StringComparison.OrdinalIgnoreCase) >= 0
                                && (text.IndexOf("доставке", StringComparison.OrdinalIgnoreCase) >= 0) || text.IndexOf("перевозке", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (documentV != null)
                                {
                                    documentV.TovarKPerevozkePrin = RemoveFirstWord(text, 4);
                                }
                                Dispatcher.Invoke(() => cb6.IsChecked = true);
                                Dispatcher.Invoke(() => AddData(11, RemoveFirstWord(text, 4)));
                                boolTovarKDostavkePrin = true;
                            }
                        }
                        if (boolPoDoverenn == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            var data = ExtractData(text);
                            if (data.powerOfAttorney != null)
                            {
                                if (data.powerOfAttorney.Length != 0)
                                {
                                    if (documentV != null)
                                    {
                                        documentV.PoDover = data.powerOfAttorney;
                                    }
                                    Dispatcher.Invoke(() => cb6.IsChecked = true);
                                    Dispatcher.Invoke(() => AddData(12, data.powerOfAttorney));
                                    boolPoDoverenn = true;
                                }
                            }
                        }
                        if (boolDoverennVidana == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            var data = ExtractData(text);
                            if (data.issuedBy != null)
                            {
                                if (data.issuedBy.Length != 0)
                                {
                                    if (documentV != null)
                                    {
                                        documentV.Vidannoi = data.issuedBy;
                                    }
                                    Dispatcher.Invoke(() => cb6.IsChecked = true);
                                    Dispatcher.Invoke(() => AddData(13, data.issuedBy));
                                    boolDoverennVidana = true;
                                }
                            }
                        }
                        if (boolTOVARNRAZDEL == false)
                        {
                            string text = string.Join("", listOfRows[i]);
                            if (text.IndexOf("ТОВАРНЫЙ", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("РАЗДЕЛ", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                //MessageBox.Show("!!!");
                                Dispatcher.Invoke(() => AddData(14, null));
                                Dispatcher.Invoke(() => cb5.IsChecked = true);
                                if (tables.Count != 0)
                                {
                                    if (tables[0] != null)
                                    {
                                        bool p = false;
                                        for (int l = 0; l < tables[0].KORDy.Count - 1; l++)
                                        {
                                            System.Drawing.Point topLeft = new System.Drawing.Point(tables[0].KORDx[0], tables[0].KORDy[l]); // Верхний левый угол
                                            System.Drawing.Point bottomRight = new System.Drawing.Point(tables[0].KORDx[1], tables[0].KORDy[l + 1]); // Нижний правый угол

                                            using (Bitmap originalImage = new Bitmap(Path.Combine(outputDirectory, $"doc1.png")))
                                            {
                                                System.Drawing.Rectangle cropArea = GetCropArea(topLeft, bottomRight);
                                                using (Bitmap croppedImage = CropImage(originalImage, cropArea))
                                                {
                                                    string tx = ExtractTextFromImage(croppedImage);
                                                    if (tx.IndexOf("№", StringComparison.OrdinalIgnoreCase) >= 0 || (tx.IndexOf("N", StringComparison.OrdinalIgnoreCase) >= 0))
                                                    {
                                                        p = true;
                                                        break;
                                                    }
                                                }
                                            }
                                        }//проверка первого столбца
                                        if (p == false)
                                        {
                                            for (int l = 2; l < tables[0].KORDy.Count - 1; l++)
                                            {
                                                DataRazdel1 razd = new DataRazdel1();
                                                string tx = null;
                                                for (int n = 0; n < tables[0].KORDx.Count - 1; n++)
                                                {
                                                    System.Drawing.Point topLeft = new System.Drawing.Point(tables[0].KORDx[n] + 5, tables[0].KORDy[l] + 5); // Верхний левый угол
                                                    System.Drawing.Point bottomRight = new System.Drawing.Point(tables[0].KORDx[n + 1] - 5, tables[0].KORDy[l + 1] - 5); // Нижний правый угол

                                                    if (tables[0].KORDx[n] == tables[0].KORDx[n + 1])
                                                    {
                                                        tables[0].KORDx.Remove(tables[0].KORDx[n]);
                                                        n--;
                                                        continue;
                                                    }

                                                    using (Bitmap originalImage = new Bitmap(Path.Combine(outputDirectory, $"doc1.png")))
                                                    {
                                                        System.Drawing.Rectangle cropArea = GetCropArea(topLeft, bottomRight);
                                                        using (Bitmap croppedImage = CropImage(originalImage, cropArea))
                                                        {
                                                            tx = ExtractTextFromImage(croppedImage);
                                                            tx = CleanString(tx);
                                                            switch (n)
                                                            {
                                                                case 0:
                                                                    razd.НаименованиеТовара = tx;
                                                                    break;
                                                                case 1:
                                                                    razd.ЕдиницаИзмерения = tx;
                                                                    break;
                                                                case 2:
                                                                    if (string.IsNullOrWhiteSpace(tx))
                                                                    {
                                                                        tx = "1";
                                                                    }
                                                                    //MessageBox.Show(tx);
                                                                    razd.Количество = tx;
                                                                    break;
                                                                case 3:
                                                                    razd.Цена = tx;
                                                                    break;
                                                                case 4:
                                                                    razd.Стоимость = tx;
                                                                    break;
                                                                case 5:
                                                                    razd.СтавкаНДС = tx;
                                                                    break;
                                                                case 6:
                                                                    razd.СуммаНДС = tx;
                                                                    break;
                                                                case 7:
                                                                    razd.СтоимостьСНДС = tx;
                                                                    break;
                                                                case 8:
                                                                    razd.Примечание = tx;
                                                                    break;
                                                            }
                                                        }
                                                    }
                                                }
                                                items.Add(razd);
                                            }
                                        }
                                        else
                                        {
                                            for (int l = 2; l < tables[0].KORDy.Count - 1; l++)
                                            {
                                                DataRazdel1 razd = new DataRazdel1();
                                                string tx = null;
                                                for (int n = 1; n < tables[0].KORDx.Count - 1; n++)
                                                {
                                                    System.Drawing.Point topLeft = new System.Drawing.Point(tables[0].KORDx[n] + 5, tables[0].KORDy[l] + 5); // Верхний левый угол
                                                    System.Drawing.Point bottomRight = new System.Drawing.Point(tables[0].KORDx[n + 1] - 5, tables[0].KORDy[l + 1] - 5); // Нижний правый угол
                                                    if (tables[0].KORDx[n] == tables[0].KORDx[n + 1])
                                                    {
                                                        tables[0].KORDx.Remove(tables[0].KORDx[n]);
                                                        n--;
                                                        continue;
                                                    }
                                                    using (Bitmap originalImage = new Bitmap(Path.Combine(outputDirectory, $"doc1.png")))
                                                    {
                                                        System.Drawing.Rectangle cropArea = GetCropArea(topLeft, bottomRight);
                                                        using (Bitmap croppedImage = CropImage(originalImage, cropArea))
                                                        {
                                                            tx = ExtractTextFromImage(croppedImage);
                                                            tx = CleanString(tx);
                                                            //MessageBox.Show(tx + "\n" + n);
                                                            switch (n)
                                                            {
                                                                case 1:
                                                                    razd.НаименованиеТовара = tx;
                                                                    break;
                                                                case 2:
                                                                    razd.ЕдиницаИзмерения = tx;
                                                                    break;
                                                                case 3:
                                                                    if (string.IsNullOrWhiteSpace(tx))
                                                                    {
                                                                        tx = "1";
                                                                    }
                                                                    //MessageBox.Show(tx);
                                                                    razd.Количество = tx;
                                                                    break;
                                                                case 4:
                                                                    razd.Цена = tx;
                                                                    break;
                                                                case 5:
                                                                    razd.Стоимость = tx;
                                                                    break;
                                                                case 6:
                                                                    razd.СтавкаНДС = tx;
                                                                    break;
                                                                case 7:
                                                                    razd.СуммаНДС = tx;
                                                                    break;
                                                                case 8:
                                                                    razd.СтоимостьСНДС = tx;
                                                                    break;
                                                                case 9:
                                                                    razd.Примечание = tx;
                                                                    break;
                                                            }
                                                        }
                                                    }
                                                }
                                                items.Add(razd);
                                            }
                                        }
                                    }
                                }
                                boolTOVARNRAZDEL = true;
                            }
                        }
                    }
                }

                //using (StreamWriter writer = new StreamWriter("M://info.txt", false, Encoding.UTF8))
                //{
                //    foreach (var row in listOfRows)
                //    {
                //        string line = string.Join("", row);
                //        writer.WriteLine(line);
                //    }
                //}
                Dispatcher.Invoke(() => progressBar.Value = 90);
                if (grid.Count != 0)
                {
                    Dispatcher.Invoke(() => buttonExcel.IsEnabled = true);
                    Dispatcher.Invoke(() => buttonXML.IsEnabled = true);
                }
                Dispatcher.Invoke(() => progressBar.Value = 100);
                if (documentV != null)
                {
                    bool checkExcel = false;
                    Dispatcher.Invoke(() => checkExcel = fastExcel.IsChecked == true);
                    if (checkExcel)
                    {
                        Dispatcher.Invoke(() => progressBar.Value = 80);
                        Dispatcher.Invoke(() => buttonExcel_Click(null, null));
                    }                    
                }
                Dispatcher.Invoke(() => progressBar.Value = 100);
                Dispatcher.Invoke(() => buttonScan.IsEnabled = true);
            });
        }
        public static string CleanString(string input)
        {
            // Удаляем указанные символы, включая переносы строк
            string cleanedString = input.Replace("\n", "")
                                        .Replace("\r", "")
                                        .Replace("|", "")
                                        .Replace("\\", "")
                                        .Replace("/", "");

            // Заменяем более одного пробела на один
            cleanedString = Regex.Replace(cleanedString, @"\s+", " ");

            // Убираем пробелы в начале и конце строки
            cleanedString = cleanedString.Trim();

            return cleanedString;
        }
        static System.Drawing.Rectangle GetCropArea(System.Drawing.Point topLeft, System.Drawing.Point bottomRight)
        {
            int width = bottomRight.X - topLeft.X;
            int height = bottomRight.Y - topLeft.Y;
            //MessageBox.Show(bottomRight.X.ToString() + "\n" + topLeft.X.ToString() + "\n" + bottomRight.Y.ToString() + "\n" + topLeft.Y.ToString());
            return new System.Drawing.Rectangle(topLeft.X, topLeft.Y, width, height);
        }
        static string ExtractTextFromImage(Bitmap image)
        {
            using (var engine = new TesseractEngine(@"D:\Programm\editor\Tesseract-OCR\tessdata", "eng+rus", EngineMode.Default))
            {
                using (var img = PixConverter.ToPix(image))
                {
                    using (var page = engine.Process(img))
                    {
                        return page.GetText();
                    }
                }
            }
        }
        static Bitmap CropImage(Bitmap original, System.Drawing.Rectangle cropArea)
        {
            Bitmap croppedImage = new Bitmap(cropArea.Width, cropArea.Height);
            using (Graphics g = Graphics.FromImage(croppedImage))
            {
                g.DrawImage(original, new System.Drawing.Rectangle(0, 0, croppedImage.Width, croppedImage.Height), cropArea, GraphicsUnit.Pixel);
            }
            string outputDirectory = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ConvertedImages");

            // Размеры нового изображения с учетом добавленных границ
            int newWidth = croppedImage.Width + 2 * 30;
            int newHeight = croppedImage.Height + 2 * 30;
            Bitmap newImage = new Bitmap(newWidth, newHeight);
            using (Graphics g = Graphics.FromImage(newImage))
            {
                g.Clear(System.Drawing.Color.White);

                // Копируем оригинальное изображение в центр нового изображения
                int x = 30;
                int y = 30;
                g.DrawImage(croppedImage, x, y, croppedImage.Width, croppedImage.Height);
            }

            newImage.Save(Path.Combine(outputDirectory, $"doc{cropArea.X}_{cropArea.Y}.png"));
            return newImage;
        }
        public void ScanTable(string currentWord, Tesseract.Rect bounds)
        {
            if (currentWord == null || currentWord.Length == 0 || currentWord == "" || currentWord == " ")
            {
                lineCoordinates.Add(Tuple.Create(bounds.X1, bounds.Y1, bounds.X2, bounds.Y2));
            }
        }
        public (string powerOfAttorney, string issuedBy) ExtractData(string input)
        {
            string powerOfAttorney = null;
            string issuedBy = null;
            string pattern1 = @"по доверенности\s*(?<powerOfAttorney>[^выданной]*)\s*выданной\s*(?<issuedBy>.*)";
            string pattern2 = @"по доверенности\s*(?<powerOfAttorney>.*)";
            string pattern3 = @"выданной\s*(?<issuedBy>.*)";
            var match = Regex.Match(input, pattern1);
            if (match.Success)
            {
                powerOfAttorney = match.Groups["powerOfAttorney"].Value.Trim();
                issuedBy = match.Groups["issuedBy"].Value.Trim();
            }
            else
            {
                match = Regex.Match(input, pattern2);
                if (match.Success)
                {
                    powerOfAttorney = match.Groups["powerOfAttorney"].Value.Trim();
                }
                match = Regex.Match(input, pattern3);
                if (match.Success)
                {
                    issuedBy = match.Groups["issuedBy"].Value.Trim();
                }
            }
            return (powerOfAttorney, issuedBy);
        }
        public string RemoveFirstWord(string input, int n = 1)
        {
            int i = 1;
            List<string> textList = input.Split().ToList(); ;
            for (int k = 0; k < textList.Count; k++)
            {
                if (textList[k] == " " || textList[k] == "")
                {
                    textList.Remove(textList[k]);
                }
            }
            for (int j = 0; j < n; j++)
            {
                i += textList[j].Length + 1;
            }
            return input.Substring(i).TrimStart();
        }
        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                var grid_ = FindParent<Grid>(button);
                if (grid_ != null)
                {
                    var stackPanel = MainStackPanel;
                    if (stackPanel != null)
                    {
                        stackPanel.Children.Remove(grid_);
                        grid.Remove(grid_);
                    }
                }
            }
            if (grid.Count == 0)
            {
                buttonExcel.IsEnabled = false;
                buttonXML.IsEnabled = false;
            }
        }
        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            DependencyObject parentObject1 = VisualTreeHelper.GetParent(parentObject);
            DependencyObject parentObject2 = VisualTreeHelper.GetParent(parentObject1);

            if (parentObject2 == null) return null;

            T parent = parentObject2 as T;
            if (parent != null)
            {
                return parent;
            }
            else
            {
                return FindParent<T>(parentObject2);
            }
        }
        private void DebugTesseractZone(string currentWord, Bitmap originalImage, Tesseract.Rect bounds, Bitmap copiedImage)
        {
            if (currentWord != null && currentWord.Length != 0 && currentWord != "" && currentWord != " " && 1 < 1)
            {
                for (int x = 0; x < originalImage.Width; x++)
                {
                    for (int y = 0; y < originalImage.Height; y++)
                    {
                        if (x >= bounds.X1 && x <= bounds.X2 && (y == bounds.Y1 || y == bounds.Y2))
                        {
                            System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                            if (pixelColor != System.Drawing.Color.Green)
                            {
                                copiedImage.SetPixel(x, y, System.Drawing.Color.Red);
                            }
                        }
                        if (y >= bounds.Y1 && y <= bounds.Y2 && (x == bounds.X1 || x == bounds.X2))
                        {
                            System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                            if (pixelColor != System.Drawing.Color.Green)
                            {
                                copiedImage.SetPixel(x, y, System.Drawing.Color.Red);
                            }
                        }
                    }
                }
            }
            if (currentWord == null || currentWord.Length == 0 || currentWord == "" || currentWord == " ")
            {
                for (int x = 0; x < originalImage.Width; x++)
                {
                    for (int y = 0; y < originalImage.Height; y++)
                    {
                        if (x >= bounds.X1 && x <= bounds.X2 && (y == bounds.Y1 || y == bounds.Y2))
                        {
                            System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                            if (pixelColor != System.Drawing.Color.Green)
                            {
                                copiedImage.SetPixel(x, y, System.Drawing.Color.Blue);
                            }
                        }
                        if (y >= bounds.Y1 && y <= bounds.Y2 && (x == bounds.X1 || x == bounds.X2))
                        {
                            System.Drawing.Color pixelColor = copiedImage.GetPixel(x, y);
                            if (pixelColor != System.Drawing.Color.Green)
                            {
                                copiedImage.SetPixel(x, y, System.Drawing.Color.Blue);
                            }
                        }
                    }
                }
            }
            copiedImage.SetPixel(bounds.X1, bounds.Y1, System.Drawing.Color.Green);
            copiedImage.SetPixel(bounds.X2, bounds.Y2, System.Drawing.Color.Green);
            //copiedImage.SetPixel(bounds.X1, bounds.Y2, System.Drawing.Color.Green);
            //copiedImage.SetPixel(bounds.X2, bounds.Y1, System.Drawing.Color.Green);
        }
        private Grid CreateDuplicatedGrid(int type, string data)
        {
            Grid grid = new Grid();
            grid.Height = 120;
            grid.Name = "dataPrefab";
            Border border = new Border();
            border.CornerRadius = new CornerRadius(15);
            border.BorderThickness = new Thickness(1);
            border.BorderBrush = System.Windows.Media.Brushes.Black;
            border.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(188, 245, 168));
            border.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
            border.Margin = new Thickness(44, 0, 44, 0);
            border.Height = 104;
            border.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            border.Name = "dataPrefabBorder";
            Grid innerGrid = new Grid();
            innerGrid.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(31, GridUnitType.Star) });
            innerGrid.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(71, GridUnitType.Star) });
            innerGrid.Name = "dataPrefabGrid";
            TextBox textBox = new TextBox();
            textBox.Name = "TextBox";
            textBox.Text = data;
            textBox.Margin = new Thickness(21, 25, 21, 17);
            textBox.TextWrapping = TextWrapping.Wrap;
            if (type == 14)
            {
                textBox.Visibility = Visibility.Hidden;
            }
            else
            {
                textBox.Visibility = Visibility.Visible;
            }
            Grid.SetRow(textBox, 1);
            ComboBox comboBox = new ComboBox();
            comboBox.Name = "comboBoxDataTypes";
            comboBox.Margin = new Thickness(22, 16, 0, 57);
            Grid.SetRowSpan(comboBox, 2);
            comboBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
            comboBox.Width = 317;
            comboBox.ItemsSource = typesData;
            comboBox.SelectedIndex = type;
            comboBox.SelectionChanged += comboBoxDataTypes_SelectionChanged;
            Button button = new Button();
            button.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
            button.VerticalAlignment = System.Windows.VerticalAlignment.Bottom;
            button.Margin = new Thickness(0, 0, 22, -15);
            button.Width = 29;
            button.Height = 29;
            button.Background = imgDelete;
            button.Style = FindResource("ImageButtonStyle") as Style;
            button.Click += DeleteButton_Click;
            Button button2 = new Button();
            button2.Name = "buttonTable";
            button2.Content = "Просмотреть таблицу";
            if (type == 14)
            {
                button2.Visibility = Visibility.Visible;
            }
            else
            {
                button2.Visibility = Visibility.Hidden;
            }
            button2.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
            button2.VerticalAlignment = System.Windows.VerticalAlignment.Top;
            button2.Margin = new Thickness(21, 25, 0, 0);
            button2.Width = 317;
            button2.Height = 29;
            button2.Click += buttonTable_Click;
            Grid.SetRow(button2, 1);
            innerGrid.Children.Add(textBox);
            innerGrid.Children.Add(comboBox);
            innerGrid.Children.Add(button);
            innerGrid.Children.Add(button2);
            border.Child = innerGrid;
            grid.Children.Add(border);
            return grid;
        }
        private void Label_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenFileButton_Click(sender, e);
        }
        private void AddDataPanel(object sender, RoutedEventArgs e)
        {
            AddData(0, "");
        }
        int nZoom = 0;
        int[] wZoom = new int[16];
        int[] hZoom = new int[16];
        private void ZoomPanel(object sender, RoutedEventArgs e)
        {
            if (nZoom == 0)
            {
                double width = imgBox.Width;
                double height = imgBox.Height;
                wZoom = new int[16];
                hZoom = new int[16];
                wZoom[0] = (int)imgBox.Width;
                hZoom[0] = (int)imgBox.Height;
                for (int i = 1; i < 16; i++)
                {
                    wZoom[i] = wZoom[i - 1] + Convert.ToInt32((float)(wZoom[i - 1] / 100) * 15);
                    hZoom[i] = hZoom[i - 1] + Convert.ToInt32((float)(hZoom[i - 1] / 100) * 15);
                }
            }
            if (nZoom < 15)
            {
                nZoom++;
                imgBox.Width = wZoom[nZoom];
                imgBox.Height = hZoom[nZoom];
            }
        }
        private void ZoomMPanel(object sender, RoutedEventArgs e)
        {
            if (nZoom >= 1)
            {
                nZoom--;
                imgBox.Width = wZoom[nZoom];
                imgBox.Height = hZoom[nZoom];
            }
        }
        private void comboBoxDataTypes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            Border parentBorder = FindParent2<Border>(comboBox);
            if (parentBorder == null) return;
            Button viewTableButton = FindChild<Button>(parentBorder, "buttonTable");
            TextBox textBox = FindChild<TextBox>(parentBorder, null);

            if (comboBox.SelectedIndex == 14)
            {
                viewTableButton.Visibility = Visibility.Visible;
                textBox.Visibility = Visibility.Hidden;
            }
            else
            {
                viewTableButton.Visibility = Visibility.Hidden;
                textBox.Visibility = Visibility.Visible;
            }
        }
        private T FindChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            if (parent == null) return null;
            T foundChild = null;

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                T childType = child as T;
                if (childType == null)
                {
                    foundChild = FindChild<T>(child, childName);
                    if (foundChild != null) break;
                }
                else if (!string.IsNullOrEmpty(childName))
                {
                    var frameworkElement = child as FrameworkElement;
                    if (frameworkElement != null && frameworkElement.Name == childName)
                    {
                        foundChild = (T)child;
                        break;
                    }
                }
                else
                {
                    foundChild = (T)child;
                    break;
                }
            }

            return foundChild;
        }
        private T FindParent2<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;

            T parent = parentObject as T;
            if (parent != null)
            {
                return parent;
            }
            else
            {
                return FindParent2<T>(parentObject);
            }
        }
        private void buttonTable_Click(object sender, RoutedEventArgs e)
        {
            if (tb == null)
            {
                tb = new Table(items);
                tb.Closed += Tb_Closed;
                tb.Show();
            }
        }
        private void Tb_Closed(object sender, EventArgs e)
        {
            tb = null;
        }
        public bool ochist = false;
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            ochist = !ochist;
            //MessageBox.Show(ochist.ToString());
        }
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            ochist = !ochist;
            //MessageBox.Show(ochist.ToString());
        }

        private void buttonExcel_Click(object sender, RoutedEventArgs e)
        {
            if (documentV != null)
            {
                documentV.table1 = items;
                documentV.ConvertToExcel(this);
                Process.Start(new ProcessStartInfo(@"Exits\file.xlsx") { UseShellExecute = true });
                Process.Start("explorer.exe", @"Exits");
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (documentV != null)
            {
                documentV.table1 = items;
                documentV.ConvertToXML(this);
            }
        }

        private async void buttonScan_Click(object sender, RoutedEventArgs e)
        {
            await Scan();
        }
    }
}
