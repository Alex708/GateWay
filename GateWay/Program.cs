using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GateWay
{
    class Program
    {
        static void Main(string[] args)
        {
            Global.Start();
            if (Global.IsFatalError)                // Если не определен файл для Паспорта отчета (журнал ошибок) ?
            {
                return;                             // -->> Выход
            }
            
            try
            {
                using (Global.writer = File.CreateText(@Global.PassportFullName))       // Открыт поток для записи совершенных действий в паспорт прогона
                {
                    Global.timer.Start();           // Запуск таймера

                    Global.OutputLine(string.Format("Стартовала программа: {0} {1}", Global.ExecProgName, Global.VersName));
                    Global.OutputLine(string.Format("Дата и время старта: {0} {1}\n", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString()));
                    Global.OutputLine(string.Format("Путь: '{0}'", Global.ExecFullName));
                    Global.OutputLine(string.Format("Пользователь: '{0}'", Environment.UserName));
                    Global.OutputLine(string.Format("Компьютер: '{0}'", Environment.MachineName));

                    Global.Initialization();        // Инициализация глобальных переменных
                    if (Global.IsFatalError)
                    {
                        Global.PrintAboutError();
                        return;                     // -->> Выход
                    }
                    //Global.mailer = new MsgMail();              // Активировать почту
                    Global.mainApp = new MainApp();

                    Global.mainApp.LoadSourceData();            // Загрузка исходных данных по Дебиторке и email адресов из файла настроек для сбора заявок
                    Global.mainApp.LoadStatusData(false);       // Загрузка статусных данных для их обновления

                    Global.mainApp.AdditionStatusData();        // Дополнение статусных данных на основе исходных данных (новые строки, актуализация email и менеджеров)

                    Global.mainApp.LoadStatusData(true);        // Перезагрузка статусных данных после их актуализации

                    Global.mainApp.ProcessDebtorData();         // Обработка данных по Дебиторке и актуализация статусов

                    Global.OutputLine("");
                    Global.timer.Stop();
                    Global.OutputLine("");
                    Global.OutputLine(string.Format("=== Общее время выполнения: {0:N3} сек", (double)Global.timer.ElapsedMilliseconds / 1000));
                    Global.OutputLine(string.Format("=== Дата в время завершения: {0} {1}\n", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString()));

                    Global.PrintAboutError();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("********** ПРОЦЕСС ПРЕРВАН!!! ОБНАРУЖЕНА ФАТАЛЬНАЯ ОШИБКА");
                Console.WriteLine("Сообщение об ошибке: '{0}'", ex.Message);
                Console.WriteLine("********** Проверьте путь к файлу Паспорта отчета в конфигурационном файле (см.конфиг)");
                Console.WriteLine(">>> Для заверения нажмите Enter");
                Console.ReadLine();
            }

            if (Global.Send && (Global.IsFatalError || Global.IsNoncriticzlError))
            {
                Global.mailer.SendPassport(Global.PassportMailAddress, Global.PassportFullName);
            }
        }
    }
}
