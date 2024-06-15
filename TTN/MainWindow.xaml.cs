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
            "ТоварКДоставкеПринял",
            "ПоДоверенности(#)",
            "ДоверенностьВыдана",
            "1ТОВАРНЫЙРАЗДЕЛ",
        };
        public List<Grid> grid = new List<Grid>();
        Table tb = null;




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
        }


        public void TesseractStart(object sender, RoutedEventArgs e)
        {
            Filter filter = new Filter();
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
                //MessageBox.Show($"Выбран файл: {selectedFilePath}");
                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(selectedFilePath, UriKind.RelativeOrAbsolute);
                bitmap.EndInit();
                imgBox.Source = bitmap;
                pathfile = selectedFilePath;
                viewbox1.Visibility = Visibility.Hidden;
                //изменение состояния кнопок
                buttonScan.IsEnabled = true;
                menuButtonScan.IsEnabled = true;
            }
        }
        public void Scan(object sender, RoutedEventArgs e)
        {
            string outputDirectory = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ConvertedImages");
            if (checkFilter.IsChecked == false)
            {
                Filter filter = new Filter();
                filter.FilterWhite(pathfile, outputDirectory);
            }

            List<string> rows = new List<string>();
            List<List<string>> listOfRows = new List<List<string>>();

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
                                    //MessageBox.Show(listOfRows.Count.ToString() + "\n" + rows.Count.ToString() + "\n" + nRow);
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

            int maxDistance = 2;

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
                                break;
                            }
                        }
                        AddData(0, data);
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
                                break;
                            }
                        }
                        AddData(1, data);
                    }
                    if (boolDateHead == false)
                    {
                        string pattern = @"\b\d{1,2}\s(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s\d{4}\b";
                        Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
                        Match match = regex.Match(string.Join("", listOfRows[i]));
                        if (match.Success)
                        {
                            AddData(3, match.Value);
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
                            AddData(4, RemoveFirstWord(text,1));
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
                            AddData(5, RemoveFirstWord(text));
                            boolGruzopoluch = true;
                        }
                    }
                    if (boolOsnovanOtpusk == false)
                    {
                        string text = string.Join("", listOfRows[i]);
                        if (text.IndexOf("Основание", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("отпуска", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AddData(6, RemoveFirstWord(text, 2));
                            boolOsnovanOtpusk = true;
                        }
                    }
                    if (boolVsegoSummaNDS == false)
                    {
                        string text = string.Join("", listOfRows[i]);
                        if (text.IndexOf("Всего", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("сумма", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("НДС", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AddData(7, RemoveFirstWord(text, 3));
                            boolVsegoSummaNDS = true;
                        }
                    }
                    if (boolVsegoStoimostSNDS == false)
                    {
                        string text = string.Join("", listOfRows[i]);
                        if (text.IndexOf("Всего", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("стоимость", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("с", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("НДС", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AddData(8, RemoveFirstWord(text, 4));
                            boolVsegoStoimostSNDS = true;
                        }
                    }
                    if (boolOtpuskRazreshil == false)
                    {
                        string text = string.Join("", listOfRows[i]);
                        if (text.IndexOf("Отпуск", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("разрешил", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AddData(9, RemoveFirstWord(text, 2));
                            boolOtpuskRazreshil = true;
                        }
                    }
                    if (boolSdalGruzootpravit == false)
                    {
                        string text = string.Join("", listOfRows[i]);
                        if (text.IndexOf("Сдал", StringComparison.OrdinalIgnoreCase) >= 0 && text.IndexOf("Грузоотправитель", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AddData(10, RemoveFirstWord(text, 2));
                            boolSdalGruzootpravit = true;
                        }
                    }
                    if (boolTovarKDostavkePrin == false)
                    {
                        string text = string.Join("", listOfRows[i]);
                        if (text.IndexOf("Товар", StringComparison.OrdinalIgnoreCase) >= 0 
                            && text.IndexOf("к", StringComparison.OrdinalIgnoreCase) >= 0 
                            && text.IndexOf("доставке", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AddData(11, RemoveFirstWord(text, 4));
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
                                AddData(12, data.powerOfAttorney);
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
                                AddData(13, data.issuedBy);
                                boolDoverennVidana = true;
                            }
                        }
                    }
                    if (boolTOVARNRAZDEL == false)
                    {
                        string text = string.Join("", listOfRows[i]);
                        if (text.IndexOf("ТОВАРНЫЙ", StringComparison.OrdinalIgnoreCase) >= 0
                            && text.IndexOf("РАЗДЕЛ", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AddData(14, null);
                            boolTOVARNRAZDEL = true;
                        }
                    }
                }
            }
            using (StreamWriter writer = new StreamWriter("M://info.txt", false, Encoding.UTF8))
            {
                foreach (var row in listOfRows)
                {
                    string line = string.Join("", row);
                    writer.WriteLine(line);
                }
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
            List<string> textList = input.Split().ToList();;
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
            if (currentWord != null && currentWord.Length != 0 && currentWord != "" && currentWord != " ")
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
            copiedImage.SetPixel(bounds.X1, bounds.Y2, System.Drawing.Color.Green);
            copiedImage.SetPixel(bounds.X2, bounds.Y1, System.Drawing.Color.Green);
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
            Grid innerGrid = new Grid();
            innerGrid.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(31, GridUnitType.Star) });
            innerGrid.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(71, GridUnitType.Star) });
            TextBox textBox = new TextBox();
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
            if(type == 14)
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
            if(nZoom == 0)
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
            if (nZoom >= 0)
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
            if(tb == null)
            {
                tb = new Table();
                tb.Closed += Tb_Closed;
                tb.Show();
            }
        }
        private void Tb_Closed(object sender, EventArgs e)
        {
            tb = null;
        }
    }
}
