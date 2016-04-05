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
    public class ManagerData
    {
        public Dictionary<string, string> ManagerDict { get; private set; }     // Коллекция с email менеджеров
        public bool IsLoaded { get; private set; }          // Признак успешной загрузки данных
        public long DownloadTime { get; private set; }      // Время загрузки данных в миллисекундах
        
        // Загрузка email по менеджерам в коллекцию (конструктор объекта)
        public ManagerData(string fileName, string sheetName)
        {
            DownloadTime = 0;
            Stopwatch timer = new Stopwatch();                      // Таймер для учета времени загрузки
            timer.Start();

            ManagerDict = new Dictionary<string, string>();         // Создание пустой коллекции

            DataTable dataTable = ExcelImport.ImportExcelXLS(fileName, sheetName);      // Загрузка исходных данных

            long skipCount = 0;         // Счетчик пропущенных записей
            long errorCount = 0;        // Счетчик ошибок класса ###
            long duplicateCount = 0;    // Счетчик дубликатов по значению 'Менеджер'
            long notEmailCount = 0;     // Счетчик строк без email адреса
            long notManagerCount = 0;   // Счетчик строк без ФИО менеджера

            IsLoaded = (dataTable != null) ? true : false;
            if (!IsLoaded)              // Если таблица пуста?
            {
                Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствуют записи (Пусто!)", Path.GetFileName(fileName), sheetName));
                Global.IsNoncriticzlError = true;
                timer.Stop();
                DownloadTime = timer.ElapsedMilliseconds;
                return;                 // -->>
            }

            long rowPos = 1;            // Текущий номер строки в Excel таблице (используется для указания на строки с ошибками)
            long rowCount = 0;          // Счетчик записей (строк) в исходной таблице
            
            string managerName;
            string emailValue;
            string emailTest;

            try
            {
                foreach (DataRow row in dataTable.Rows)     // Просмотр исходных данных
                {
                    if (row[0] != DBNull.Value && (string)row[0] == "Итого")        // Если конец таблицы ?
                    {
                        break;
                    }

                    rowPos++;               // Номер текущей строки (для диагностики)
                    rowCount++;             // Счетчик обработанных строк

                    if (row[0] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В таблице '{0}({1})' отсутствует значение 'Менеджер' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        notManagerCount++;
                        continue;                           // -->>   пропустить эту запись
                    }

                    if (row[1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В таблице '{0}({1})' отсутствует значение 'Почта' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        notEmailCount++;
                        continue;                           // -->>   пропустить эту запись
                    }

                    managerName = (string)row[0];           // ФИО менеджера
                    emailValue = (string)row[1];            // Email менеджера

                    // Формирование нового элемента коллекции
                    
                    if (!ManagerDict.TryGetValue(managerName, out emailTest))       // Если нет запись с таким ключем ?
                    {
                        ManagerDict.Add(managerName, emailValue);                   // Добавление новой записи в коллекцию
                    }
                    else
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' обнаружен дубликан по значению 'Менеджер' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        duplicateCount++;
                        continue;           // -->>   пропустить эту запись
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
                return;         // Выход -->> 
            }

            Global.OutputLine(string.Format("- Общее количество записей в исходной таблице '{0} ({1})': {2}", Path.GetFileName(@fileName), sheetName, rowCount));

            Global.OutputLine(string.Format("- Количество пропущенных записей: {0}", skipCount));
            
            if (errorCount > 0)
            {
                Global.OutputLine(string.Format("- Количество ошибок класса ###: {0}", errorCount));
            }
            if (duplicateCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей с дубликатами по значению 'Менеджер': {0}", duplicateCount));
            }
            if (notManagerCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей без менеджера: {0}", notManagerCount));
            }
            if (notEmailCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей без email: {0}", notEmailCount));
            }

            timer.Stop();
            DownloadTime = timer.ElapsedMilliseconds;
        }
    }
}
