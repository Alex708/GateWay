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
    public class OrdEmlData
    {
        public Dictionary<string, string> OrdEmlDict { get; private set; }     // Коллекция email адресов по клиентам (клиент -> email)
        public bool IsLoaded { get; private set; }          // Признак успешной загрузки данных
        public long DownloadTime { get; private set; }      // Время загрузки данных в миллисекундах

        // Загрузка данных по email клиентов в коллекцию (конструктор объекта)
        public OrdEmlData(string fileName, string sheetName)
        {
            DownloadTime = 0;
            Stopwatch timer = new Stopwatch();              // Таймер для учета времени загрузки
            timer.Start();

            OrdEmlDict = new Dictionary<string, string>();   // Создание пустой коллекции email адресов по клиентам
            
            string clientName;
            string oldEmailValue;
            string newEmailValue;

            DataTable dataTable = ExcelImport.ImportExcelXLS(fileName, sheetName);    // Загрузка данных из excel
            IsLoaded = (dataTable != null) ? true : false;
            if (!IsLoaded)
            {
                Global.OutputLine(string.Format("*** Ошибка! В файле '{0}({1})' отсутствуют записи (Пусто!)", Path.GetFileName(fileName), sheetName));
                Global.IsFatalError = true;
                timer.Stop();
                return;  // Выход -->> 
            }

            long count = 0;             // Счетчик записей
            long skipCount = 0;         // Счетчик пропущенных строк (проигнорированных)
            long errorCount = 0;        // Счетчик ошибок класса ###
            long duplicateCount = 0;    // Счетчик дубликатов по значению 'Клиент'

            try
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    count++;

                    if (row[0] != DBNull.Value && (string)row[0] == "Итог")     // Если встретилась итоговая строка умной таблицы?
                    {
                        break;      // Конец таблицы        // -->>
                    }

                    if (row[0] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Клиент' в строке '{2}'", Path.GetFileName(fileName), sheetName, count + 1));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        continue;                           // -->>
                    }
                    else if (((string)row[0]).Contains('"'))
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' присутствует кавычка в значении 'Клиент' в строке '{2}'", Path.GetFileName(fileName), sheetName, count + 1));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        continue;                           // -->>
                    }

                    if (row[4] == DBNull.Value)             // Если нет email адресов ?     (запись не пропускается!!!)
                    {
                        newEmailValue = null;
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Почта' (нет e-mail адресов) в строке '{2}'", Path.GetFileName(fileName), sheetName, count + 1));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        continue;                           // -->>
                    }
                    else
                    {
                        newEmailValue = (string)row[4];     // Email адреса
                    }

                    clientName = (string)row[0];            // Клиент (ключ поиска в коллекции)

                    if (!OrdEmlDict.TryGetValue(clientName, out oldEmailValue))     // Если не найдено? (т.е. нет дубликата)
                    {
                        OrdEmlDict.Add(clientName, newEmailValue);                  // Добавление записи в коллекцию email адресов по клиентам
                    }
                    else
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' обнаружен дубликан по значению 'Клиент' в строке '{2}'", Path.GetFileName(fileName), sheetName, count + 1));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        duplicateCount++;
                        continue;                       // -->>
                    }
                }
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("*** Ошибка! При попытке загрузить данные из Excel файла '{0}' (лист: '{1}', строка: '{2}') возникла ошибка '{3}'", Path.GetFullPath(@fileName), sheetName, count + 1, ex.Message));
                Global.OutputLine(string.Format("Проверьте ВСЮ таблицу перед очередным запуском!"));
                Global.IsFatalError = true;
                IsLoaded = false;
                return;  // Выход -->> 
            }

            Global.OutputLine(string.Format("- Общее количество записей в исходной таблице '{0}({1})': {2}", Path.GetFileName(fileName), sheetName, count));
            
            if (skipCount > 0)
            {
                Global.OutputLine(string.Format("- Количество пропущенных записей: {0}", skipCount));
            }
            if (errorCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей с ошибоками класса ###: {0}", errorCount));
            }
            if (duplicateCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей с дубликатами по значению 'Клиент': {0}", duplicateCount));
            }
            
            timer.Stop();
            DownloadTime = timer.ElapsedMilliseconds;
        }
    }

}
