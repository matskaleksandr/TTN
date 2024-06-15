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
        public Table()
        {
            InitializeComponent();
            var items = new List<DataRazdel1>
            {
                new DataRazdel1 { НаименованиеТовара = "Кабель", 
                    ЕдиницаИзмерения = "м", 
                    Количество = 100 ,
                    Цена = 100,
                    Стоимость = 100,
                    СтавкаНДС = "15%",
                    СуммаНДС = 100,
                    СтоимостьСНДС = 100,
                    Примечание = "Он очень крутой"},
            };
            dataGrid.ItemsSource = items;
        }

        public class DataRazdel1
        {
            public string НаименованиеТовара { get; set; }
            public string ЕдиницаИзмерения { get; set; }
            public int Количество { get; set; }
            public double Цена { get; set; }
            public double Стоимость { get; set; }
            public string СтавкаНДС { get; set; }
            public double СуммаНДС { get; set; }
            public double СтоимостьСНДС { get; set; }
            public string Примечание { get; set; }
        }
    }
}
