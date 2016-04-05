using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GateWay
{
    // Класс для представления записи по дебиторки
    public class DbtDblItem
    {
        public string ClientCode { get; set; }              // Код клиента
        public string ClientName { get; set; }              // Название клиента
        public double RubTotSum { get; set; }               // Сумма оплаты в рублях (полная задолженность)
        public double RubDueSum { get; set; }               // Сумма оплаты в рублях (просроченная задолженность)
        public DateTime? PayDueDate { get; set; }           // Дата оплаты
        public string ManagerName { get; set; }             // Менеджер
        public string UnitName { get; set; }                // Бизнес-единица
        public double UsdTotSum { get; set; }               // Сумма оплаты в USD (полная задолженность)
        public double EurTotSum { get; set; }               // Сумма оплаты в Euro (полная задолженность)
        public double UsdDueSum { get; set; }               // Сумма оплаты в USD (просроченная задолженность)
        public double EurDueSum { get; set; }               // Сумма оплаты в Euro (просроченная задолженность)
        public int Count { get; set; }                      // Счетчик записей

        public DbtDblItem()
        {
            ClientCode = null;
            ClientName = null;
            RubTotSum = 0;
            RubDueSum = 0;
            PayDueDate = null;
            ManagerName = null;
            UnitName = null;
            UsdTotSum = 0;
            EurTotSum = 0;
            UsdDueSum = 0;
            EurDueSum = 0;
            Count = 0;
        }
    }

    // Класс для представления данных по клиенту
    public class DbtClnItem
    {
        public string ClientCode { get; set; }              // Код клиента
        public string ClientName { get; set; }              // Название клиента
        public string ManagerName { get; set; }             // Менеджер

        public DbtClnItem()
        {
            ClientCode = null;
            ClientName = null;
            ManagerName = null;
        }
    }

    // Класс для формирования коллекции данных по клиентам
    public class DebtorData
    {
        public Dictionary<string, DbtDblItem> DbtDtlDict { get; private set; }    // Коллекция с данными по дебиторке (ключ: Код клиента + БЕ)
        public Dictionary<string, DbtClnItem> DbtClnDict { get; private set; }         // Коллекция с данными по клиентам (ключ: Код клиента)

        public bool IsLoaded { get; private set; }              // Признак успешной загрузки данных
        public long DownloadTime { get; private set; }          // Время загрузки данных в миллисекундах

        public long InputRowTotCount { get; private set; }      // Количество записей учтенных в коллекции (полная задолженность)
        public long InputRowDueCount { get; private set; }      // Количество записей учтенных в коллекции (просроченная задолженность)

        public double TotalRubTotSum { get; private set; }      // Общая сумма полной задолженности в рублях
        public double TotalUsdTotSum { get; private set; }      // Общая сумма полной задолженности в USD
        public double TotalEurTotSum { get; private set; }      // Общая сумма полной задолженности в EUR

        public double TotalRubDueSum { get; private set; }      // Общая сумма просроченной задолженности в рублях
        public double TotalUsdDueSum { get; private set; }      // Общая сумма просроченной задолженности в USD
        public double TotalEurDueSum { get; private set; }      // Общая сумма просроченной задолженности в EUR

        // Загрузка данных по дебиторке в коллекцию (конструктор объекта)
        public DebtorData(string fileName, string sheetName)
        {
            DownloadTime = 0;
            Stopwatch timer = new Stopwatch();                  // Таймер для учета времени загрузки
            timer.Start();

            DataTable dataTable = ExcelImport.ImportExcelXLS(fileName, sheetName);    // Загрузка исходных данных из выгрузки по просроченной дебиторке
            IsLoaded = (dataTable != null) ? true : false;
            if (!IsLoaded)
            {
                Global.OutputLine(string.Format("*** Ошибка! В файле '{0}({1})' отсутствуют записи (Пусто!)", Path.GetFileName(fileName), sheetName));
                Global.IsFatalError = true;
                timer.Stop();
                return;  // Выход -->> 
            }

            DbtDtlDict = new Dictionary<string, DbtDblItem>();      // Создание пустой коллекции для данных по дебиторке
            DbtClnDict = new Dictionary<string, DbtClnItem>();      // Создание пустой коллекции для данных по клиентам

            DbtDblItem newItem;         // Для формирования нового элемента коллекции по дебиторке
            DbtDblItem oldItem;         // Для обновления существующего элемента коллекции по дебиторке

            DbtClnItem newItem_;        // Для формирования нового элемента коллекции по клиенту
            DbtClnItem oldItem_;        // Для обновления существующего элемента коллекции по клиенту
           
            string checkCode;           // Для проверки Кода клиента
            string checkGroup;          // Для проверки Группы клиентов
            string checkValuta;         // Для проверки Валюты
            string checkUnit;           // Для проверки БЕ

            long expiredDays;           // Просрочено в днях (если >= 0)

            string keyItem;             // Для формирования значения ключа элемента коллекции

            InputRowTotCount = 0;       // Счетчик записей исходной таблицы, которые были учтены в колекции (полная задолженность)
            InputRowDueCount = 0;       // Счетчик записей исходной таблицы, которые были учтены в колекции (просроченная задолженность) 

            TotalRubTotSum = 0;         // Общая сумма полной задолженности в Рублях
            TotalUsdTotSum = 0;         // Общая сумма полной задолженности в USD
            TotalEurTotSum = 0;         // Общая сумма полной задолженности в EUR

            TotalRubDueSum = 0;         // Общая сумма просроченной задолженности в Рублях
            TotalUsdDueSum = 0;         // Общая сумма просроченной задолженности в USD
            TotalEurDueSum = 0;         // Общая сумма просроченной задолженности в EUR

            long rowPos = 1;            // Текущий номер строки в Excel таблице (используется для указания на строки с ошибками)
            long rowCount = 0;          // Счетчик строк в исходной таблице

            long errorCount = 0;        // Счетчик пропущенных строк из-за ошибок ###
            long skipCount = 0;         // Счетчик пропущенных строк по заданным условиям (не Кл, не КлиентыРос, не просрочено)

            bool isError;

            try
            {
                foreach (DataRow row in dataTable.Rows)     // Просмотр исходных данных
                {
                    rowPos++;               // Номер текущей строки (для диагностики)

                    if (rowPos <= 3)
                    {
                        continue;           // Пропустить 2 строки над таблицей     -->>
                    }

                    rowCount++;             // Счетчик обработанных строк

                    isError = false;

                    if (row[Global._ClientCodePos - 1] == DBNull.Value)     // Если нет Кода клиента?
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Код клиент' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        continue;                                           // -->>
                    }

                    try
                    {
                        checkCode = (string)row[Global._ClientCodePos - 1]; // Преобразование в строку! Здесь контролируется форматирование. Если не текст, то будет исключение!
                        if (checkCode.Substring(0, 2) != "Кл")              // Код клиента должен иметь префикс "Кл", иначе пропускаем строку (не обрабатываем!)
                        {
                            skipCount++;
                            continue;                                       // -->>
                        }
                    }
                    catch (Exception ex)
                    {
                        Global.OutputLine(string.Format("*** Ошибка! В файле '{0}({1})' неправильный ФОРМАТ в 1-й колонке (Код клиента), в строке {2}. Сообщение: '{3}'", Path.GetFileName(fileName), sheetName, rowPos, ex.Message));
                        Global.OutputLine(string.Format("Для исправления формата нужно сделать 'Пробный отчет' с помощью файла 'Дебиторка.xmlx'"));
                        Global.IsFatalError = true;
                        break;                                              // -->>
                    }

                    if (row[Global._GroupNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Группа клиентов' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        continue;                                           // -->>
                    }

                    checkGroup = (string)row[Global._GroupNamePos - 1];
                    if (checkGroup == Global._GroupRosSkip || checkGroup == Global._ClientExSkip)   // Строки с этими кодами не обрабатывать!
                    {
                        skipCount++;
                        continue;                                           // -->>
                    }

                    checkUnit = (string)row[Global._UnitNamePos - 1];
                    if (checkUnit == Global._UnitSkip)                      // Строки с этими кодами не обрабатывать (Геркулес Питер)!
                    {
                        skipCount++;
                        continue;                                           // -->>
                    }

                    if (row[Global._ExpiredDaysPos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Просрочено' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        continue;                                           // -->>
                    }

                    if (row[Global._ClientNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Название клиента' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }
                    else if (((string)row[Global._ClientNamePos - 1]).Contains('"'))
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' присутствует кавычка в значении 'Название клиента' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }

                    if (row[Global._PaymentSumPos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Сумма к оплате' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }

                    if (row[Global._PaymentDatePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение даты 'Оплатить' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }

                    if (row[Global._ManagerNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Менеджер' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }

                    if (row[Global._UnitNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'БЕ' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }

                    if (row[Global._ValutaSumPos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Сумма в валюте' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }

                    if (row[Global._ValutaNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Валюта' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        isError = true;
                    }

                    if (isError)            // Если была ошибка ### отсутствия значений ?
                    {
                        errorCount++;
                        continue;           // -->>
                    }

                    expiredDays = (long)(double)row[Global._ExpiredDaysPos - 1];        // Сколько просрочено в днях

                    // Формирование нового элемента коллекции
                    newItem = new DbtDblItem();
                    newItem.ClientCode = (string)row[Global._ClientCodePos - 1];        // Код клиента
                    newItem.ClientName = (string)row[Global._ClientNamePos - 1];        // Название клиента
                    newItem.RubTotSum = (double)row[Global._PaymentSumPos - 1];         // Сумма к оплате в рублях (полная задолженность)
                    if (expiredDays >= 0)       // Если просрочено ?
                    {
                        newItem.RubDueSum = (double)row[Global._PaymentSumPos - 1];     // Сумма к оплате в рублях (просроченная задолженность)
                    }
                    newItem.PayDueDate = (DateTime)row[Global._PaymentDatePos - 1];     // Дата оплаты
                    newItem.ManagerName = (string)row[Global._ManagerNamePos - 1];      // ФИО Менеджера
                    newItem.UnitName = (string)row[Global._UnitNamePos - 1];            // Имя БЕ
                    checkValuta = (string)row[Global._ValutaNamePos - 1];               // Валюта
                    if (checkValuta == Global._USD)
                    {
                        newItem.UsdTotSum = (double)row[Global._ValutaSumPos - 1];      // Сумма к оплате в USD (полная задолженность)
                        if (expiredDays >= 0)       // Если просрочено ?
                        {
                            newItem.UsdDueSum = (double)row[Global._ValutaSumPos - 1];  // Сумма к оплате в USD (просроченная задолженность)
                        }
                    }
                    else if (checkValuta == Global._EUR)
                    {
                        newItem.EurTotSum = (double)row[Global._ValutaSumPos - 1];      // Сумма к оплате в EUR (полная задолженность)
                        if (expiredDays >= 0)       // Если просрочено ?
                        {
                            newItem.EurDueSum = (double)row[Global._ValutaSumPos - 1];  // Сумма к оплате в EUR (просроченная задолженность)
                        }
                    }

                    TotalRubTotSum += newItem.RubTotSum;                    // Общая сумма полной задолженности в Рублях
                    TotalUsdTotSum += newItem.UsdTotSum;                    // Общая сумма полной задолженности в USD
                    TotalEurTotSum += newItem.EurTotSum;                    // Общая сумма полной задолженности в EUR

                    if (expiredDays >= 0)           // Если просрочено ?
                    {
                        TotalRubDueSum += newItem.RubDueSum;                // Общая сумма просроченной задолженности в Рублях
                        TotalUsdDueSum += newItem.UsdDueSum;                // Общая сумма просроченной задолженности в USD
                        TotalEurDueSum += newItem.EurDueSum;                // Общая сумма просроченной задолженности в EUR
                    }

                    keyItem = newItem.ClientCode + newItem.UnitName;        // КЛЮЧ поиска в коллекции
                    
                    if (DbtDtlDict.TryGetValue(keyItem, out oldItem))       // Если есть запись с таким ключем ?
                    {
                        oldItem.RubTotSum += newItem.RubTotSum;             // Подсуммирование (полная задолженность)
                        oldItem.UsdTotSum += newItem.UsdTotSum;
                        oldItem.EurTotSum += newItem.EurTotSum;
                        oldItem.RubDueSum += newItem.RubDueSum;             // Подсуммирование (просроченная задолженность)
                        oldItem.UsdDueSum += newItem.UsdDueSum;
                        oldItem.EurDueSum += newItem.EurDueSum;
                        oldItem.Count++;
                        DbtDtlDict[keyItem] = oldItem;                      // Замена записи в коллекции (по ключу)
                    }
                    else
                    {
                        newItem.Count = 1;
                        DbtDtlDict.Add(keyItem, newItem);                   // Добавление новой записи
                    }

                    InputRowTotCount++;                                     // Количество записей учтенных в коллекции (полная задолженность)
                    if (expiredDays >= 0)           // Если просрочено ?
                    {
                        InputRowDueCount++;                                 // Количество записей учтенных в коллекции (просроченная задолженность)
                    }

                    // Формирование нового элемента для дополнительной коллекции по клиентам
                    newItem_ = new DbtClnItem();
                    newItem_.ClientCode = (string)row[Global._ClientCodePos - 1];       // Код клиента
                    newItem_.ClientName = (string)row[Global._ClientNamePos - 1];       // Название клиента
                    newItem_.ManagerName = (string)row[Global._ManagerNamePos - 1];     // ФИО Менеджера

                    if (!DbtClnDict.TryGetValue(newItem_.ClientCode, out oldItem_))     // Если нет записи с таким ключем ?
                    {
                        DbtClnDict.Add(newItem_.ClientCode, newItem_);                  // Добавление новой записи в дополнительную коллекцию
                    }
                }
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("*** Ошибка! При попытке загрузить данные из Excel файла '{0}' (лист: '{1}', строка: '{2}') возникла ошибка '{3}'", Path.GetFileName(@fileName), sheetName, rowPos, ex.Message));
                Global.OutputLine(string.Format("Проверьте ВСЮ таблицу перед очередным запуском!"));
                Global.IsFatalError = true;
                IsLoaded = false;
                timer.Stop();
                return;  // Выход -->> 
            }

            Global.OutputLine(string.Format("- Общее количество записей в исходной таблице '{0} ({1})': {2}", Path.GetFileName(@fileName), sheetName, rowCount));
            Global.OutputLine(string.Format("- Количество записей с полной задолженностью: {0}", InputRowTotCount));
            Global.OutputLine(string.Format("- Количество записей с просроченной задолженностью: {0}", InputRowDueCount));
            if (skipCount > 0)
            {
                Global.OutputLine(string.Format("- Количество пропущенных записей по стандартным условиям: {0}", skipCount));
            }
            if (errorCount > 0)
            {
                Global.OutputLine(string.Format("- Количество пропущенных записей из-за ошибок класса ###: {0}", errorCount));
            }

            Global.OutputLine(string.Format("- Общая сумма полной задолженности в Рублях: {0:N}", TotalRubTotSum));
            Global.OutputLine(string.Format("- Общая сумма полной задолженности в USD: {0:N}", TotalUsdTotSum));
            Global.OutputLine(string.Format("- Общая сумма полной задолженности в EUR: {0:N}", TotalEurTotSum));

            Global.OutputLine(string.Format("- Общая сумма просроченной задолженности в Рублях: {0:N}", TotalRubDueSum));
            Global.OutputLine(string.Format("- Общая сумма просроченной задолженности в USD: {0:N}", TotalUsdDueSum));
            Global.OutputLine(string.Format("- Общая сумма просроченной задолженности в EUR: {0:N}", TotalEurDueSum));

            Global.OutputLine(string.Format("- Итоговое количество строк в коллекции дебиторки: {0}", DbtDtlDict.Count));

            timer.Stop();
            DownloadTime = timer.ElapsedMilliseconds;
        }
    }
}
