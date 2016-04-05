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
    // Класс для представления записи с данными БЕ
    public class UnitItem
    {
        public string UnitName { get; set; }                // Обозначение БЕ
        public string UnitTitle { get; set; }               // Полное официальное название компании БЕ (юр.лицо)
        public string InnCode { get; set; }                 // Код ИНН
        public string AccCode { get; set; }                 // Расчетный счет
        public string BikCode { get; set; }                 // БИК и К/С
        public string UnitEmail { get; set; }               // Почта БО БЕ
        public string UnitManager { get; set; }             // ФИО ответственного сотрудника БО БЕ
        public string StampFileName { get; set; }           // Имя файла с изображением печати и подписи
 
        public UnitItem()
        {
            UnitName = null;
            UnitTitle = null;
            InnCode = null;
            AccCode = null;
            BikCode = null;
            UnitEmail = null;
            UnitManager = null;
            StampFileName = null;
        }
    }

    public class UnitData
    {
        public Dictionary<string, UnitItem> UnitDict { get; private set; }      // Коллекция с данными по БЕ ((ключ: БЕ)
        public bool IsLoaded { get; private set; }          // Признак успешной загрузки данных
        public long DownloadTime { get; private set; }      // Время загрузки данных в миллисекундах

         // Загрузка данных по БЕ (конструктор объекта)
        public UnitData(string fileName, string sheetName)
        {
            DownloadTime = 0;
            Stopwatch timer = new Stopwatch();              // Таймер для учета времени загрузки
            timer.Start();

            UnitDict = new Dictionary<string, UnitItem>();                              // Создание пустой коллекции

            DataTable dataTable = ExcelImport.ImportExcelXLS(fileName, sheetName);      // Загрузка исходных данных
            IsLoaded = (dataTable != null) ? true : false;

            if (!IsLoaded)              // Если таблица пуста?
            {
                Global.OutputLine(string.Format("В файле '{0}({1})' отсутствуют записи (Пусто!)", Path.GetFileName(fileName), sheetName));
                timer.Stop();
                DownloadTime = timer.ElapsedMilliseconds;
                return;                 // -->>
            }

            UnitItem newItem;           // Для формирования нового элемента коллекции
            UnitItem oldItem;           // Для обновления существующего элемента коллекции

            string keyItem;             // Для формирования значения ключа элемента коллекции

            long rowPos = 1;            // Текущий номер строки в Excel таблице (используется для указания на строки с ошибками)
            long rowCount = 0;          // Счетчик записей (строк) в исходной таблице

            long errorCount = 0;        // Счетчик ошибок класса ###
            long skipCount = 0;         // Счетчик пропущенных строк
            long duplicateCount = 0;    // Счетчик дубликатов по значению 'БЕ'

            bool isError;

            try
            {
                foreach (DataRow row in dataTable.Rows)     // Просмотр исходных данных для загрузки в коллекцию
                {
                    if (row[Global._untUntNamePos - 1] == DBNull.Value ||
                        row[Global._untUntNamePos - 1] != DBNull.Value &&
                        (string)row[Global._untUntNamePos] == "Итого")      // Если конец таблицы?
                    {
                        break;
                    }

                    rowPos++;               // Номер текущей строки (для диагностики)
                    rowCount++;             // Счетчик обработанных строк

                    isError = false;

                    if (row[Global._untUntTitlePos - 1] == DBNull.Value)             // Если нет полного названия БЕ?
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Название БЕ' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._untInnCodePos - 1] == DBNull.Value)             // Если нет Кода ИНН?
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'ИНН КПП' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._untAccCodePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Расчетный счет' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._untBikCodePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'БИК к/с' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._untUntMailPos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Почта БО' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._untUntManagerPos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'ФИО сотрудника БО' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (row[Global._untStampFileNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Печать и подпись' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        errorCount++;
                        isError = true;
                    }

                    if (isError)        // Если была ошибка ### отсутствия значений ?
                    {
                        skipCount++;
                        continue;                               // -->>   пропустить эту запись
                    }

                    // Формирование нового элемента коллекции
                    newItem = new UnitItem();
                    newItem.UnitName = (string)row[Global._untUntNamePos - 1];              // Обозначение БЕ
                    newItem.UnitTitle = (string)row[Global._untUntTitlePos - 1];            // Полное название БЕ
                    newItem.InnCode = (string)row[Global._untInnCodePos - 1];               // ИНН
                    newItem.AccCode = (string)row[Global._untAccCodePos - 1];               // Расчетный счет
                    newItem.BikCode = (string)row[Global._untBikCodePos - 1];               // БИК
                    newItem.UnitEmail = (string)row[Global._untUntMailPos - 1];             // Почта БО
                    newItem.UnitManager = (string)row[Global._untUntManagerPos - 1];        // ФИО сотрудника БО
                    newItem.StampFileName = (string)row[Global._untStampFileNamePos - 1];   // Имя файла с изображением печати и подписи
                   
                    keyItem = newItem.UnitName;                             // КЛЮЧ поиска в коллекции

                    if (!UnitDict.TryGetValue(keyItem, out oldItem))        // Если нет запись с таким ключем ?
                    {
                        UnitDict.Add(keyItem, newItem);                     // Добавление новой записи в коллекцию
                    }
                    else
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' обнаружен дубликан по значению 'БЕ' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
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
                Global.OutputLine(string.Format("- Количество записей с дубликатами по значению 'БЕ': {0}", duplicateCount));
            }
            Global.OutputLine(string.Format("- Итоговое количество записей в коллекции: {0}", UnitDict.Count));

            timer.Stop();
            DownloadTime = timer.ElapsedMilliseconds;
        }
    }
}
