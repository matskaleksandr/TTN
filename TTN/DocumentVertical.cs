using Patagames.Pdf.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TTN
{
    internal class DocumentVertical
    {
        //данные
        public string GruzOtpr;                                  //УНП грузоотправитель
        public string GruzPoluch;                                //УНП грузополучатель

        public string Date;                                      //Дата

        public string Avto;                                      //Автомобиль
        public string Pricep;                                    //Прицеп

        public string KPutList;                                  //К путевому листу номер

        public string Voditel;                                   //Водитель

        public string ZakazchikPerevozki;                        //Заказчик перевозки

        public string GruzOtprName;                              //Грузоотправитель название

        public string GruzPoluchName;                            //Грузоотправитель название

        public string OsnOtpusk;                                 //Основание отпуска

        public string PunktPogruzk;                              //Пункт погрузки
        public string PunktRazgruzki;                            //Пункт разгрузки

        public string Pereadresovka;                             //Переадресовка

        //Разделы
        public bool TovarnRazdel;

        public bool PogruzRazgruz;

        public bool ProchSved;

        //данные
        public string VsegoSummNDS;                              //Всего сумма НДС

        public string VsegoStoimSNDS;                            //Всего стоимость с НДС

        public string VsegoMassGruz;                             //Всего масса груза

        public string OtpuskRazresh;                             //Отпуск разрешил

        public string SdalGruzootpav;                            //Сдал грузоотправитель

        public string NoPlomb;                                   //Номер пломбы

        public string VsegoKolGruzMest;                          //Всего количество грузовых мест

        public string TovarKPerevozkePrin;                       //Товар к перевозке принял

        public string PoDover;                                   //По доверенности

        public string Vidannoi;                                  //Выданной

        public string PrinGruzopoluch;                           //Принял грузополучатель

        public string NoPlomb2;                                  //Номер промбы
    }
}
