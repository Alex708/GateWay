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
    // Класс для представления записи по статусам
    public class StsStsItem
    {
        public string UnitName { get; set; }                // БЕ
        public string ClientCode { get; set; }              // Код клиента
        public string ClientName { get; set; }              // Название клиента
        public double RubTotSum { get; set; }               // Сумма оплаты в рублях (полная задолженность)
        public double RubDueSum { get; set; }               // Сумма оплаты в рублях (просроченная задолженность)
        public double UsdTotSum { get; set; }               // Сумма оплаты в USD (полная задолженность)
        public double EurTotSum { get; set; }               // Сумма оплаты в Euro (полная задолженность)
        public double UsdDueSum { get; set; }               // Сумма оплаты в USD (просроченная задолженность)
        public double EurDueSum { get; set; }               // Сумма оплаты в Euro (просроченная задолженность)
        public DateTime? date1 { get; set; }                // Дата выдачи предупреждения
        public DateTime? date2 { get; set; }                // Дата выдачи запрета на отгрузки
        public long RowNum { get; set; }                    // Номер строки в Excel таблице

        public StsStsItem()
        {
            UnitName = null;
            ClientCode = null;
            ClientName = null;
            RubTotSum = 0;
            RubDueSum = 0;
            UsdTotSum = 0;
            EurTotSum = 0;
            UsdDueSum = 0;
            EurDueSum = 0;
            date1 = null;
            date2 = null;
            RowNum = 0;
        }
    }

    public class StsStsData
    {
        public Dictionary<string, StsStsItem> StsStsDict { get; private set; }     // Коллекция с данными по статусам ((ключ: Код клиента + БЕ)
        public bool IsLoaded { get; private set; }          // Признак успешной загрузки данных
        public long DownloadTime { get; private set; }      // Время загрузки данных в миллисекундах

          // Загрузка данных по статусам в коллекцию (конструктор объекта)
        public StsStsData(string fileName, string sheetName, bool isProc)
        {
            DownloadTime = 0;
            Stopwatch timer = new Stopwatch();              // Таймер для учета времени загрузки
            timer.Start();

            StsStsDict = new Dictionary<string, StsStsItem>();                          // Создание пустой коллекции

            DataTable dataTable = ExcelImport.ImportExcelXLS(fileName, sheetName);      // Загрузка исходных данных
            IsLoaded = (dataTable != null) ? true : false;

            if (!IsLoaded)              // Если таблица пуста?
            {
                Global.OutputLine(string.Format("В файле '{0}({1})' отсутствуют записи (Пусто!)", Path.GetFileName(fileName), sheetName));
                timer.Stop();
                DownloadTime = timer.ElapsedMilliseconds;
                return;                 // -->>
            }

            StsStsItem newItem;         // Для формирования нового элемента коллекции
            StsStsItem oldItem;         // Для обновления существующего элемента коллекции

            DateTime? checkDate1;       // Для проверки Дата1
            DateTime? checkDate2;       // Для проверки Дата2

            string keyItem;             // Для формирования значения ключа элемента коллекции

            long rowPos = 1;            // Текущий номер строки в Excel таблице (используется для указания на строки с ошибками)
            long rowCount = 0;          // Счетчик записей (строк) в исходной таблице

            long errorCount = 0;        // Счетчик ошибок класса ###
            long skipCount = 0;         // Счетчик пропущенных строк
            long duplicateCount = 0;    // Счетчик дубликатов по значению 'Код клиента'

            bool isError;

            try
            {
                foreach (DataRow row in dataTable.Rows)     // Просмотр исходных данных для загрузки в коллекцию
                {
                    if (row[Global._stsUntNamePos - 1] == DBNull.Value || 
                        row[Global._stsUntNamePos - 1] != DBNull.Value && 
                        (string)row[Global._stsUntNamePos - 1] == "Итого")              // Если конец таблицы?
                    {
                        break;
                    }

                    rowPos++;               // Номер текущей строки (для диагностики)
                    rowCount++;             // Счетчик обработанных строк

                    isError = false;

                    if (row[Global._stsClnCodePos - 1] == DBNull.Value)                 // Если нет Кода клиента?
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Код клиент' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._stsUntNamePos - 1] == DBNull.Value)                 // Если нет БЕ?
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'БЕ' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._stsClnNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Название клиента' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }
                    else if (((string)row[Global._stsClnNamePos - 1]).Contains('"'))
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' присутствует кавычка в значении 'Название клиента' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    checkDate1 = null;
                    if (row[Global._stsData1Pos - 1] != DBNull.Value)                   // Если в поле Дата1 не пусто?
                    {
                        try
                        {
                            checkDate1 = (DateTime)row[Global._stsData1Pos - 1];        // Преобразование значение в дату! Если не дата, то будет исключение!
                        }
                        catch (Exception ex)
                        {
                            Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' некорректное значение Дата1, в строке {2}. Сообщение: '{3}'", Path.GetFileName(fileName), sheetName, rowPos, ex.Message));
                            Global.IsNoncriticzlError = true;
                            errorCount++;
                            isError = true;
                        }
                    }

                    checkDate2 = null;
                    if (row[Global._stsData2Pos - 1] != DBNull.Value)                   // Если в поле Дата2 не пусто?
                    {
                        try
                        {
                            checkDate2 = (DateTime)row[Global._stsData2Pos - 1];        // Преобразование значение в дату! Если не дата, то будет исключение!
                        }
                        catch (Exception ex)
                        {
                            Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' некорректное значение Дата2, в строке {2}. Сообщение: '{3}'", Path.GetFileName(fileName), sheetName, rowPos, ex.Message));
                            Global.IsNoncriticzlError = true;
                            errorCount++;
                            isError = true;
                        }
                    }

                    if (isError)        // Если была ошибка ### отсутствия значений ?
                    {
                        skipCount++;
                        continue;       // -->>   пропустить эту запись
                    }

                    // Формирование нового элемента коллекции
                    newItem = new StsStsItem();
                    newItem.UnitName = (string)row[Global._stsUntNamePos - 1];      // БЕ
                    newItem.ClientCode = (string)row[Global._stsClnCodePos - 1];    // Код клиента
                    newItem.ClientName = (string)row[Global._stsClnNamePos - 1];    // Название клиента
                    if (row[Global._stsRubTotSumPos - 1] != DBNull.Value)
                    {
                        newItem.RubTotSum = (double)row[Global._stsRubTotSumPos - 1];
                    }
                    if (row[Global._stsEurTotSumPos - 1] != DBNull.Value)
                    {
                        newItem.EurTotSum = (double)row[Global._stsEurTotSumPos - 1];
                    }
                    if (row[Global._stsUsdTotSumPos - 1] != DBNull.Value)
                    {
                        newItem.UsdTotSum = (double)row[Global._stsUsdTotSumPos - 1];
                    }
                    if (row[Global._stsRubDueSumPos - 1] != DBNull.Value)
                    {
                        newItem.RubDueSum = (double)row[Global._stsRubDueSumPos - 1];
                    }
                    if (row[Global._stsEurDueSumPos - 1] != DBNull.Value)
                    {
                        newItem.EurDueSum = (double)row[Global._stsEurDueSumPos - 1];
                    }
                    if (row[Global._stsUsdDueSumPos - 1] != DBNull.Value)
                    {
                        newItem.UsdDueSum = (double)row[Global._stsUsdDueSumPos - 1];
                    }
                    newItem.date1 = checkDate1;                                     // Дата предупреждения
                    newItem.date2 = checkDate2;                                     // Дата запрета
                    newItem.RowNum = rowPos;                                        // Номер строки на Excel листе (для привязки к Excel листу!!!)

                    keyItem = newItem.ClientCode + newItem.UnitName;                // КЛЮЧ поиска в коллекции
                    
                    if (!StsStsDict.TryGetValue(keyItem, out oldItem))              // Если нет запись с таким ключем ?
                    {
                        StsStsDict.Add(keyItem, newItem);                           // Добавление новой записи в коллекцию
                    }
                    else
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' обнаружен дубликан по значению 'Код клиента+БЕ' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        duplicateCount++;
                        errorCount++;
                        skipCount++;
                        continue;                               // -->>
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
            
            Global.OutputLine(string.Format("- Количество пропущенных (проигнорированных) записей: {0}", skipCount));
           
            if (errorCount > 0)
            {
                Global.OutputLine(string.Format("- Количество ошибок класса ###: {0}", errorCount));
            }
            if (duplicateCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей с дубликатами по значению 'Код клиента': {0}", duplicateCount));
            }
            Global.OutputLine(string.Format("- Итоговое количество записей в коллекции: {0}", StsStsDict.Count));

            timer.Stop();
            DownloadTime = timer.ElapsedMilliseconds;
        }
    }
}
