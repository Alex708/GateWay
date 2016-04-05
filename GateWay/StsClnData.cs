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
    public class StsClnItem
    {
        public string ClientCode { get; set; }              // Код клиента
        public string ClientName { get; set; }              // Название клиента
        public string ManagerName { get; set; }             // Менеджер ответственный за клиента
        public string EmailValue { get; set; }              // Email адреса клиента
        public long RowNum { get; set; }                    // Номер строки в Excel таблице

        public StsClnItem()
        {
            ClientCode = null;
            ClientName = null;
            ManagerName = null;
            EmailValue = null;
            RowNum = 0;
        }
    }

    public class StsClnData
    {
        public Dictionary<string, StsClnItem> StsClnDict { get; private set; }     // Коллекция с данными по клиенту
        public bool IsLoaded { get; private set; }          // Признак успешной загрузки данных
        public long DownloadTime { get; private set; }      // Время загрузки данных в миллисекундах

        // Загрузка данных по клиентам в коллекцию (конструктор объекта)
        public StsClnData(string fileName, string sheetName, bool isProc)
        {
            DownloadTime = 0;
            Stopwatch timer = new Stopwatch();                      // Таймер для учета времени загрузки
            timer.Start();

            StsClnDict = new Dictionary<string, StsClnItem>();      // Создание пустой коллекции

            DataTable dataTable = ExcelImport.ImportExcelXLS(fileName, sheetName);    // Загрузка исходных данных

            IsLoaded = (dataTable != null) ? true : false;
            if (!IsLoaded)              // Если таблица пуста?
            {
                Global.OutputLine(string.Format("В файле '{0}({1})' отсутствуют записи (Пусто!)", Path.GetFileName(fileName), sheetName));
                timer.Stop();
                DownloadTime = timer.ElapsedMilliseconds;
                return;                 // -->>
            }

            StsClnItem newItem;         // Для формирования нового элемента коллекции
            StsClnItem oldItem;         // Для обновления существующего элемента коллекции

            string keyItem;             // Для формирования значения ключа элемента коллекции

            long rowPos = 1;            // Текущий номер строки в Excel таблице (используется для указания на строки с ошибками)
            long rowCount = 0;          // Счетчик записей (строк) в исходной таблице

            long skipCount = 0;         // Счетчик пропущенных строк (отсутствуют важные данные)
            long errorCount = 0;        // Счетчик ошибок типа ###
            long duplicateCount = 0;    // Счетчик дубликатов по значению 'Код клиента'

            long notEmailCount = 0;     // Счетчик строк без email адресов
            long notManagerCount = 0;   // Счетчик строк без ФИО менеджера

            long notSendCount = 0;      // Счетчик строк с клиентами, по которым не работать (не рассылать)

            string managerName;
            string emailValue;

            try
            {
                foreach (DataRow row in dataTable.Rows)     // Просмотр исходных данных
                {
                    if (row[Global._clnClnCodePos - 1] == DBNull.Value || 
                        row[Global._clnClnCodePos - 1] != DBNull.Value && 
                        (string)row[Global._clnClnCodePos - 1] == "Итого")  // Если конец таблицы ?
                    {
                        break;
                    }

                    rowPos++;               // Номер текущей строки (для диагностики)
                    rowCount++;             // Счетчик обработанных строк

                    if (isProc && 
                        row[Global._clnNoSendPos - 1] != DBNull.Value && (int)(double)row[Global._clnNoSendPos - 1] == 1)       // Если не отсылать?
                    {
                        skipCount++;
                        notSendCount++;
                        continue;                               // -->>   пропустить эту запись
                    }

                    if (row[Global._clnClnNamePos - 1] == DBNull.Value)
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Название клиента' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        skipCount++;
                        errorCount++;
                        continue;                               // -->>   пропустить эту запись
                    }

                    if (row[Global._clnMngNamePos - 1] == DBNull.Value)
                    {
                        managerName = null;
                        if (isProc)
                        {
                            Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Менеджер' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                            Global.IsNoncriticzlError = true;
                            errorCount++;
                        }
                        notManagerCount++;
                    }
                    else
                    {
                        managerName = (string)row[Global._clnMngNamePos - 1];       // ФИО Менеджера
                    }

                    if (row[Global._clnEmailPos - 1] == DBNull.Value)
                    {
                        emailValue = null;
                        if (isProc)
                        {
                            Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' отсутствует значение 'Почта' (email адреса клиента) в строке {2}", Path.GetFileName(fileName), sheetName, rowPos));
                            Global.IsNoncriticzlError = true;
                            errorCount++;
                        }
                        notEmailCount++;
                    }
                    else
                    {
                        emailValue = (string)row[Global._clnEmailPos - 1];          // Email адреса клиента
                    }

                    // Формирование нового элемента коллекции
                    newItem = new StsClnItem();                 // Создание новой записи для добавления в коллекцию
                    newItem.ClientCode = (string)row[Global._clnClnCodePos - 1];    // Код клиента
                    newItem.ClientName = (string)row[Global._clnClnNamePos - 1];    // Название клиента
                    newItem.ManagerName = managerName;                              // ФИО Менеджера
                    newItem.EmailValue = emailValue;                                // Email адреса клиента
                    newItem.RowNum = rowPos;                                        // Номер строки на Excel листе (для привязки к Excel листу!!!)

                    keyItem = newItem.ClientCode;                                   // Формирование ключа поиска в коллекции

                    if (!StsClnDict.TryGetValue(keyItem, out oldItem))              // Если нет запись с таким ключем ?
                    {
                        StsClnDict.Add(keyItem, newItem);                           // Добавление новой записи в коллекцию
                    }
                    else
                    {
                        Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})' обнаружен дубликан по значению 'Код клиента' в строке '{2}'", Path.GetFileName(fileName), sheetName, rowPos));
                        Global.IsNoncriticzlError = true;
                        skipCount++;        // Счетчик пропущенных (проигнорированных) строк
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
            
            Global.OutputLine(string.Format("- Количество пропущенных записей (проигнорированных) : {0}", skipCount));
            
            if (notSendCount > 0)
            {
                Global.OutputLine(string.Format("- Количество строк, по которым не отсылать: {0}", notSendCount));
            }
            if (errorCount > 0)
            {
                Global.OutputLine(string.Format("- Количество ошибок класса ###: {0}", errorCount));
            }

            if (notManagerCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей без менеджера: {0}", notManagerCount));
            }
            if (notEmailCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей без email: {0}", notEmailCount));
            }
            if (duplicateCount > 0)
            {
                Global.OutputLine(string.Format("- Количество записей с дубликатами по значению 'Код клиента': {0}", duplicateCount));
            }
            Global.OutputLine(string.Format("- Итоговое количество записей в коллекции: {0}", StsClnDict.Count));

            timer.Stop();
            DownloadTime = timer.ElapsedMilliseconds;
        }
    }
}
