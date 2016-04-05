using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace GateWay
{
    public enum Traffic { Green, Yellow, Red }

    public static class Global
    {
        public readonly static int WindowHeight = 40;
        public readonly static int WindowWidth = 95;

        public const string VersName = "[версия 1.0  от 31.03.2016]";

        public static string ExecFullName { get; private set; }
        public static string ExecProgName { get; private set; }

        public static string WinTitleName { get; private set; }

        public static bool IsFatalError { get; set; }                   // Глобальный признак фатальной ошибки
        public static bool IsNoncriticzlError { get; set; }             // Глобальный признак некритичной ошибки

        public static string PassportFolderName { get; set; }           // Путь к папке с файлами Паспортов прогона
        public static string PassportFullName { get; set; }             // Полное имя файла для сохранения Паспорта текущего прогона
        public static string PassportMailAddress { get; set; }          // Email адреса для пересылки Паспорта прогона

        public static string FileName_Debetorka { get; set; }           // Имя файла и исходными данными по Дебиторке (ежедневная выгрузка)
        public static string FileName_OrdersSetup { get; set; }         // Имя настроечного файла из проекта "Заявки"
        public static string FileName_DebtorStatus { get; set; }        // Имя рабочего файла для статусов Клиентов в разрезе БЕ

        public static string FolderName_Template { get; set; }          // Имя папки для шаблонов извещений
        public static string FolderName_Notice { get; set; }            // Имя папки для извещений

        public static string FileName_TemplateNotice1 { get; set; }     // Имя папки для извещений 1 (Предупреждение о возможной приостановки отгрузок)
        public static string FileName_TemplateNotice2 { get; set; }     // Имя папки для извещений 2 (Извещение о приостановке отгрузок)
        public static string FileName_TemplateNotice3 { get; set; }     // Имя папки для извещений 3 (Извещение о возобновлении отгрузок)

        public static MsgMail mailer { get; set; }                      // Outlook mail
        public static MainApp mainApp { get; set; }                     // Main Application

        public static bool Visible { get; set; }                        // Признак визуализации (отчет после создания показывается и файл с отчетом не закрывается)
        public static bool Send { get; set; }                           // Призран рассылки по e-mail

        public readonly static Color DefaultRowColor = Color.FromArgb(212, 250, 252);       // Цвет для подкраски строк по умолчанию
        public readonly static Color SkipRowColor = Color.FromArgb(207, 207, 207);          // Цвет для подкраски пропущенных строк
        public readonly static Color AddedRowColor = Color.FromArgb(150, 150, 240);         // Цвет для подкраски добавленных строк
        public readonly static Color UpdatedCellColor = Color.FromArgb(131, 252, 240);      // Цвет для подкраски обновленных ячеек
        public readonly static Color TrafficGreenColor = Color.FromArgb(144, 238, 144);     // Цвет светофора Зеленый
        public readonly static Color TrafficYellowColor = Color.FromArgb(255, 246, 143);    // Цвет светофора Желтый
        public readonly static Color TrafficRedColor = Color.FromArgb(255, 105, 106);       // Цвет светофора Красный

        // Позиции колонок в таблице "savDebetorkaReport1.xlsx"
        public const int _ClientCodePos = 1;                            // Номер колонки Код клиента
        public const int _ClientNamePos = 2;                            // Номер колонки Название клиента
        public const int _PaymentSumPos = 6;                            // Номер колонки Сумма к оплате (в рублях)
        public const int _PaymentDatePos = 7;                           // Номер колонки Оплатить до (дата)
        public const int _ExpiredDaysPos = 8;                           // Номер колонки Просрочено
        public const int _ManagerNamePos = 13;                          // Номер колонки Менеджер
        public const int _GroupNamePos = 14;                            // Номер колонки Группа клиентов
        public const int _UnitNamePos = 15;                             // Номер колонки БЕ
        public const int _ValutaSumPos = 17;                            // Номер колонки Сумма в валюте
        public const int _ValutaNamePos = 18;                           // Номер колонки Валюта

        public const string _Debetorka = "Акт выверки";

        public const string _Statuses = "Статусы";
        public const string _Clients = "Клиенты";
        public const string _Managers = "Менеджеры";
        public const string _Units = "БЕ";

        public const string _GroupRosSkip = "ГруппаРос";                // Пропускать (не обрабатывать!)
        public const string _ClientExSkip = "КлиентыЭкс";               // Пропускать (не обрабатывать!)

        public const string _UnitSkip = "grp";                          // Пропускать (не обрабатывать!)

        public const string _USD = "USD";
        public const string _EUR = "EUR";

        // Позиции колонок в таблице "DebtorStatus.xlsx(Клиенты)"
        public const int _clnClnCodePos = 1;
        public const int _clnClnNamePos = 2;
        public const int _clnMngNamePos = 3;
        public const int _clnEmailPos = 4;                              // Email адреса клиентов
        public const int _clnNoSendPos = 5;                             // Признак того, что по этому клиенту не работаем (если 1)
        public const int _clnUserPos = 6;
        public const int _clnAddTimePos = 7;
        public const int _clnUpdTimePos = 8;

        // Позиции колонок в таблице "DebtorStatus.xlsx(Статусы)"
        public const int _stsUntNamePos = 1;
        public const int _stsClnCodePos = 2;
        public const int _stsClnNamePos = 3;
        public const int _stsRubTotSumPos = 4;                          // Номер колонки с полной задолженностью в Рублях
        public const int _stsEurTotSumPos = 5;
        public const int _stsUsdTotSumPos = 6;
        public const int _stsRubDueSumPos = 7;                          // Номер колонки с просроченной задолженностью в Рублях
        public const int _stsEurDueSumPos = 8;
        public const int _stsUsdDueSumPos = 9;
        public const int _stsData1Pos = 10;                             // Номер колонки для дат выдачи предупреждений                      
        public const int _stsData2Pos = 11;                             // Номер колонки для дат приостановки отгрузок
        public const int _stsUserPos = 12;
        public const int _stsAddTimePos = 13;
        public const int _stsUpdTimePos = 14;

        // Позиции колонок в таблице "DebtorStatus.xlsx(БЕ)"
        public const int _untUntNamePos = 1;
        public const int _untUntTitlePos = 2;                           // Номер колонки с полным названием БЕ
        public const int _untInnCodePos = 3;                            // Номер колонки с ИНН и КПП 
        public const int _untAccCodePos = 4;                            // Номер колонки с расчетным счетом 
        public const int _untBikCodePos = 5;
        public const int _untUntMailPos = 6;                            // Номер колонки с email БО БЕ
        public const int _untUntManagerPos = 7;
        public const int _untStampFileNamePos = 8;                      // Номер колонки для имен файлов с изображениями печати и подписи

        public const double MIN_SUM = 1000;                             // Пороговое значение просроченной задолженности, выше которого выдается предупреждение

        public const int DEAD_LINE = 4;                                 // Период после предупреждения, в течение которого просроченная задолженность должна быть погашена

        public static Stopwatch timer = new Stopwatch();                // Таймер

        public static StreamWriter writer;                              // Поток для вывода паспорта прогона в текстовый файл

        public static List<string> passport = new List<string>();       // Паспорт прогона

        public static void OutputLine(string item, bool OnConsole = true)
        {
            if (OnConsole)
            {
                Console.WriteLine(item);
            }
            writer.WriteLine(item);
            passport.Add(item);
        }

        public static string TimeSpanString(TimeSpan time)
        {
            return string.Format("{0,4}:{1:00}", time.Days * 24 + time.Hours, time.Minutes);
        }

        public static string ComplexKey(int positionNum, string productName)
        {
            return string.Format("{0,1} {1}", positionNum.ToString(), productName);
        }

        static Global()
        {
            IsFatalError = false;
            IsNoncriticzlError = false;

            PassportFolderName = null;
            PassportFullName = null; 
            PassportMailAddress = null;

            FileName_Debetorka = null;
            FileName_OrdersSetup = null;
            FileName_DebtorStatus = null;

            FolderName_Template = null;
            FolderName_Notice = null;

            FileName_TemplateNotice1 = null;
            FileName_TemplateNotice2 = null;
            FileName_TemplateNotice3 = null;

            int windowWidth;
            int windowHeight;

            bool IsConsoleSize = false;         // Признак определенности размеров консоли

            ExecFullName = System.Reflection.Assembly.GetExecutingAssembly().Location;      // Полное имя исполняемого файла
            ExecProgName = Path.GetFileNameWithoutExtension(ExecFullName);                  // Краткое имя исполняемого файла (без расширения)

            WinTitleName = ExecProgName + " " + VersName;                                   // Имя программы с версией (для заголовка)

            Console.Title = WinTitleName;

            // Определение размеров консоли
            try
            {
                if (ConfigurationManager.AppSettings.AllKeys.Contains("WindowWidth"))
                {
                    windowWidth = int.Parse(ConfigurationManager.AppSettings["WindowWidth"]);
                }
                else
                {
                    windowWidth = 100;
                }
                if (ConfigurationManager.AppSettings.AllKeys.Contains("WindowHeight"))
                {
                    windowHeight = int.Parse(ConfigurationManager.AppSettings["WindowHeight"]);
                }
                else
                {
                    windowHeight = 40;
                }
                Console.WindowWidth = windowWidth;
                Console.WindowHeight = windowHeight;
                IsConsoleSize = true;
            }
            catch
            {
                IsConsoleSize = false;
            }

            if (!IsConsoleSize)     // Если размеры консоли не установлены ?
            {
                Console.WindowHeight = Console.LargestWindowHeight - 20;
                Console.WindowWidth = Console.LargestWindowWidth - 30;
                Console.WriteLine(string.Format("### Размеры окна консоли подобраны по разрешению экрана: WindowWidth = '{0}, WindowHeight = {1}'", Console.WindowWidth, Console.WindowHeight));
                Console.WriteLine();
                Global.IsNoncriticzlError = true;
            }

            Console.ForegroundColor = ConsoleColor.Green;       // Цвет текста в консоли
        }

        public static void Start()
        {
            if (ConfigurationManager.AppSettings.AllKeys.Contains("PassportFolderName"))
            {
                Global.PassportFolderName = ConfigurationManager.AppSettings["PassportFolderName"];
            }
            else
            {
                IsFatalError = true;
                Console.WriteLine("********** Не определен путь к папке для Паспортов прогона (см. в конфиг PassportFolderName)");
            }

            if (IsFatalError)
            {
                Console.WriteLine("********** ПРОЦЕСС ПРЕРВАН!!! ОБНАРУЖЕНА ФАТАЛЬНАЯ ОШИБКА !!!");
                Console.WriteLine("********** Проверьте весь конфигурационный файл и исправьте обнаруженный ошибки");
                Console.WriteLine(">>> Для заверения нажмите Enter");
                Console.ReadLine();
            }
            else
            {
                string timeStamp = (DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss")).Replace(':', '-');
                Global.PassportFullName = Path.Combine(Global.PassportFolderName, "LOG_" + Global.ExecProgName) + " (" + timeStamp + ")" + ".txt";  // Полное имя файла для поспорта прогона
            }
        }

        // Проверка наличия имени файла в App.config
        private static void CheckConfig(string fileName, out string fullFileName)
        {
            fullFileName = null;
            if (ConfigurationManager.AppSettings.AllKeys.Contains(fileName))
            {
                fullFileName = ConfigurationManager.AppSettings[fileName];
            }
            else
            {
                IsFatalError = true;
                Global.OutputLine(string.Format("*** Ошибка! Не определен путь к файлу '{0}' (см.конфиг)", fileName));
            }
        }

        public static void Initialization()
        {
            string fullFileName;
            try
            {
                fullFileName = null;
                CheckConfig("FileName_Debetorka", out fullFileName);        // Имя файла с исходными данными по просроченной дебиторке
                Global.FileName_Debetorka = fullFileName;

                fullFileName = null;
                CheckConfig("FileName_DebtorStatus", out fullFileName);     // Имя файла со статусами клиентов по БЕ
                Global.FileName_DebtorStatus = fullFileName;

                fullFileName = null;
                CheckConfig("FileName_OrdersSetup", out fullFileName);      // Имя файла со статусами клиентов по БЕ
                Global.FileName_OrdersSetup = fullFileName;

                if (ConfigurationManager.AppSettings.AllKeys.Contains("FolderName_Template"))
                {
                    Global.FolderName_Template = ConfigurationManager.AppSettings["FolderName_Template"];
                    if (!Directory.Exists(Global.FolderName_Template))
                    {
                        IsFatalError = true;
                        Global.OutputLine(string.Format("*** Ошибка! Не найдена папка для шаблонов извещений: {0}", Global.FolderName_Template));
                    }
                }
                else
                {
                    IsFatalError = true;
                    Global.OutputLine(string.Format("*** Ошибка! Не определен путь к папке для шаблонов извещений FolderName_Template (см.конфиг)"));
                }

                if (ConfigurationManager.AppSettings.AllKeys.Contains("FolderName_Notice"))
                {
                    Global.FolderName_Notice = ConfigurationManager.AppSettings["FolderName_Notice"];
                    if (!Directory.Exists(Global.FolderName_Notice))
                    {
                        IsFatalError = true;
                        Global.OutputLine(string.Format("*** Ошибка! Не найдена папка для извещений: {0}", Global.FolderName_Notice));
                    }
                }
                else
                {
                    IsFatalError = true;
                    Global.OutputLine(string.Format("*** Ошибка! Не определен путь к папке извещений FolderName_Notice (см.конфиг)"));
                }

                if (ConfigurationManager.AppSettings.AllKeys.Contains("FileName_TemplateNotice1"))
                {
                    Global.FileName_TemplateNotice1 = ConfigurationManager.AppSettings["FileName_TemplateNotice1"];
                    if (!File.Exists(Path.Combine(Global.FolderName_Template, Global.FileName_TemplateNotice1)))
                    {
                        IsFatalError = true;
                        Global.OutputLine(string.Format("*** Ошибка! Не найден файл с шаблоном извещения 1: {0}", Global.FileName_TemplateNotice1));
                    }
                }
                else
                {
                    IsFatalError = true;
                    Global.OutputLine(string.Format("*** Ошибка! Не определено имя файла с шаблоном извещения FileName_TemplateNotice1 (см.конфиг)"));
                }

                if (ConfigurationManager.AppSettings.AllKeys.Contains("FileName_TemplateNotice2"))
                {
                    Global.FileName_TemplateNotice2 = ConfigurationManager.AppSettings["FileName_TemplateNotice2"];
                    if (!File.Exists(Path.Combine(Global.FolderName_Template, Global.FileName_TemplateNotice2)))
                    {
                        IsFatalError = true;
                        Global.OutputLine(string.Format("*** Ошибка! Не найден файл с шаблоном извещения 2: {0}", Global.FileName_TemplateNotice2));
                    }
                }
                else
                {
                    IsFatalError = true;
                    Global.OutputLine(string.Format("*** Ошибка! Не определено имя файла с шаблоном извещения FileName_TemplateNotice2 (см.конфиг)"));
                }

                if (ConfigurationManager.AppSettings.AllKeys.Contains("FileName_TemplateNotice3"))
                {
                    Global.FileName_TemplateNotice3 = ConfigurationManager.AppSettings["FileName_TemplateNotice3"];
                    if (!File.Exists(Path.Combine(Global.FolderName_Template, Global.FileName_TemplateNotice3)))
                    {
                        IsFatalError = true;
                        Global.OutputLine(string.Format("*** Ошибка! Не найден файл с шаблоном извещения 3: {0}", Global.FileName_TemplateNotice3));
                    }
                }
                else
                {
                    IsFatalError = true;
                    Global.OutputLine(string.Format("*** Ошибка! Не определено имя файла с шаблоном извещения FileName_TemplateNotice3 (см.конфиг)"));
                }

                if (ConfigurationManager.AppSettings.AllKeys.Contains("PassportMailAddress"))
                {
                    Global.PassportMailAddress = ConfigurationManager.AppSettings["PassportMailAddress"];
                }
                else
                {
                    IsFatalError = true;
                    Global.OutputLine(string.Format("*** Ошибка! Не определены email адреса для рассылки паспорта прогона PassportMailAddress (см.конфиг)"));
                }

                if (ConfigurationManager.AppSettings.AllKeys.Contains("Visible"))
                {
                    Global.Visible = (ConfigurationManager.AppSettings["Visible"] == "1") ? true : false;
                }
                else
                {
                    IsNoncriticzlError = true;
                    Global.OutputLine(string.Format("### ошибка! Не определен параметр Visible (см.конфиг)"));
                    Global.Visible = true;
                }

                if (ConfigurationManager.AppSettings.AllKeys.Contains("Send"))
                {
                    Global.Send = (ConfigurationManager.AppSettings["Send"] == "1") ? true : false;
                }
                else
                {
                    IsNoncriticzlError = true;
                    Global.OutputLine(string.Format("### ошибка! Не определен параметр Send (см.конфиг)"));
                    Global.Send = true;
                }
            }
            catch (Exception ex)
            {
                IsFatalError = true;
                Global.OutputLine(string.Format("*** Ошибка! При обработке конфигурационного файла 'App.config' произошел сбой (см.конфиг): '{0}'", ex.Message));
            }
        }

        // Вывод общей информации об ошибках
        public static void PrintAboutError()
        {
            if (Global.IsFatalError)                // Если есть фатальные ошибки при создании и/или отправке Оперативного отчета НКУ ?
            {
                Global.OutputLine(string.Format("********** ПРОЦЕСС ПРЕРВАН!!! ОБНАРУЖЕНЫ ФАТАЛЬНЫЕ ОШИБКИ!!!"));
                Global.OutputLine(string.Format("********** УСТРАНИТЕ ОШИБКИ И ПОВТОРИТЕ СНОВА"));
            }
            else if (Global.IsNoncriticzlError)     // Если есть некритичные ошибки при создании Оперативного отчета НКУ ?
            {
                Global.OutputLine(string.Format("########## ОБНАРУЖЕНЫ НЕКРИТИЧНЫЕ ОШИБКИ!"));
                Global.OutputLine(string.Format("########## УСТРАНИТЕ ОШИБКИ ДО СЛЕДУЮЩИЕГО ЗАПУСКА!"));
            }
            else
            {
                Global.OutputLine(string.Format("Прогон выполнен УСПЕШНО (без ошибок)!"));
            }
        }

        //public static DateTime FirstDateOfMonth(DateTime date)
        //{
        //    return new DateTime(date.Year, date.Month, 1);
        //}
    }
}
