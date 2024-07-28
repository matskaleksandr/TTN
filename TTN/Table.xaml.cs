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
using System.Windows.Shapes;

namespace TTN
{
    /// <summary>
    /// Логика взаимодействия для Table.xaml
    /// </summary>
    public partial class Table : Window
    {
        List<DataRazdel1> items_ = new List<DataRazdel1>();
        MainWindow window = null;

        public Table(List<DataRazdel1> items, MainWindow wind)
        {
            window = wind;
            InitializeComponent();
            dataGrid.ItemsSource = items;
        }

        public class DataRazdel1
        {
            public string НаименованиеТовара { get; set; }
            public string ЕдиницаИзмерения { get; set; }
            public string Количество { get; set; }
            public string Цена { get; set; }
            public string Стоимость { get; set; }
            public string СтавкаНДС { get; set; }
            public string СуммаНДС { get; set; }
            public string СтоимостьСНДС { get; set; }
            public string Примечание { get; set; }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            items_ = dataGrid.ItemsSource.Cast<DataRazdel1>().ToList();
            window.items = items_;
        }
    }
}
