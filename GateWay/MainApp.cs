using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace GateWay
{
    public class MainApp
    {
        private Excel.Application excelApp { get; set; }    // Excel приложение
        private Word._Application wordApp { get; set; }     // Word приложение

        private DebtorData debtorData { get; set; }         // Данные по дебиторке из savDebetorkaReport1 (в 2 коллекции)
        private OrdEmlData ordEmlData { get; set; }         // Email адреса клиентов из OrdersSetup
        private StsClnData stsClnData { get; set; }         // Данные по клиентам из DebtorStatus
        private StsStsData stsStsData { get; set; }         // Данные по статусам из DebtorStatus
        private ManagerData managerData { get; set; }       // Данные по менеджерам из DebtorStatus
        private UnitData unitData { get; set; }             // Данные по БЕ из DebtorStatus

        private long warningMsgCount { get; set; }          // Счетчик предупреждений о возможной приостановке отгрузок
        private long eraseMsgCount { get; set; }            // Счетчик сброса даты выдачи предупреждения
        private long stopMsgCount { get; set; }             // Счетчик извещений о приостановке отгрузок
        private long startMsgCount { get; set; }            // Счетчик извещений о возобновлении отгрузок

        private bool isUpdateRow { get; set; }
        private DateTime currDate { get; set; }             // Текущая дата

        ~MainApp()
        {
            try
            {
                excelApp.Quit();
                excelApp = null;
                wordApp.Quit();
                wordApp = null;
            }
            finally { }
        }

        public MainApp()
        {
            debtorData = null;
            ordEmlData = null;
            stsClnData = null;
            stsStsData = null;
            managerData = null;
            unitData = null;
            excelApp = new Excel.Application();             // Создание Excel приложения
            wordApp = new Word.Application();               // Создание Word приложения
        }

        // **********************************************************************
        // Загрузка исходных данных (из 2-х таблиц)
        public void LoadSourceData()
        {
            debtorData = new DebtorData(Global.FileName_Debetorka, Global._Debetorka);                  // Загрузка данных просроченной дебиторки
            ordEmlData = new OrdEmlData(Global.FileName_OrdersSetup, Global._Clients);                  // Загрузка данных email клиентов из файла настроек приложения для сбора Заявок клиентов
        }

        // **********************************************************************
        // Загрузка статусных данных (из 2-х таблиц)
        public void LoadStatusData(bool isProc)
        // isProc == false - предварительная загрузка для дополнения статусных данных
        // isProc == true  - загрузка для актуализации статусных данных и рассылки писем
        {
            this.stsClnData = new StsClnData(Global.FileName_DebtorStatus, Global._Clients, isProc);    // Загрузка данных по клиентам (менеджеры, email клиентов)
            this.stsStsData = new StsStsData(Global.FileName_DebtorStatus, Global._Statuses, isProc);   // Загрузка главной таблицы статусных данных
            if (isProc)
            {
                this.managerData = new ManagerData(Global.FileName_DebtorStatus, Global._Managers);     // Загрузка данных email менеджеров
                this.unitData = new UnitData(Global.FileName_DebtorStatus, Global._Units);              // Загрузка данных БЕ
            }
        }

        // **********************************************************************
        // ДОПОЛНЕНИЕ СТАТУСНЫЙ ДАННЫХ на основе исходных данных
        public void AdditionStatusData()
        {
            Global.OutputLine("");
            Global.OutputLine(string.Format("==>> ДОПОЛНЕНИЕ СТАТУСНЫХ ДАННЫХ"));

            Excel.Workbook book = null;
            Excel.Worksheet sheet = null;

            try
            {
                book = excelApp.Workbooks.Open(Global.FileName_DebtorStatus, Editable: true);   // Открытие файла со статусными данными
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("*** Ошибка: Не удалось открыть файл: '{0}'. Сообщение: '{1}'", Global.FileName_DebtorStatus, ex.Message));
                Global.IsFatalError = true;
                return;                 // -->>
            }

            Global.OutputLine("");
            Global.OutputLine(string.Format("--> Дополнение данных по Клиентам (менеджеры, email клиентов"));
            
            sheet = (Excel.Worksheet)book.Worksheets[Global._Clients];                          // Лист КЛИЕНТЫ
            
            StsClnItem clientItem;
            string emailValue;

            long startAddPos;           // Номер строки в таблице
            long addCount;              // Счетчик добавленных записей
            long updCount;              // Счетчик обновленных записей

            bool isUpdate;

            // Поиск позиции для добавления первой новой строки на лист Клиенты
            for (startAddPos = 2; sheet.Cells[startAddPos, 1].Value != null; startAddPos++) { }
            Global.OutputLine(string.Format("Номер стартовой строки для добавления новых строк в таблицу '{0}({1})': {2}",  Path.GetFileName(Global.FileName_DebtorStatus), Global._Clients, startAddPos));
           
            addCount = 0;
            updCount = 0;

            foreach (DbtClnItem item in debtorData.DbtClnDict.Values)       // Просмотр дополнительной коллекции с данными по клиенту
            {
                isUpdate = false;
                if (stsClnData.StsClnDict.TryGetValue(item.ClientCode, out clientItem))                 // Если найден клиент (по коду) в клиентской базе ?
                {
                    if (!Exist(clientItem.ManagerName))                                                 // Если в базе клиентов менеджер не существует?
                    {
                        if (Exist(item.ManagerName))                                                    // Если менеджер существует в дебиторке?
                        {
                            UpdateCellValue(sheet, clientItem.RowNum, Global._clnMngNamePos, item.ManagerName);     // Определить менеджера!
                            isUpdate = true;
                        }
                    }
                    else if (Exist(item.ManagerName) && clientItem.ManagerName != item.ManagerName)     // Если менеджер изменился?
                    {
                        UpdateCellValue(sheet, clientItem.RowNum, Global._clnMngNamePos, item.ManagerName);        // Изменить менеджера
                        isUpdate = true;
                    }

                    if (ordEmlData.OrdEmlDict.TryGetValue(item.ClientName, out emailValue) && Exist(emailValue))    // Если найдена запись с email адресами клиета?
                    {
                        if (!Exist(clientItem.EmailValue) || emailValue != clientItem.EmailValue)                   // Если в клиентской базе нет email или в дебиторке другой email?
                        {
                            UpdateCellValue(sheet, clientItem.RowNum, Global._clnEmailPos, emailValue);             // Обновить email адреса
                            isUpdate = true;
                        }
                    }
                    if (isUpdate)
                    {
                        updCount++;
                    }
                }
                else        // не найден, значит добавить новую строку в таблицу!
                {
                    sheet.Cells[startAddPos, Global._clnClnCodePos] = item.ClientCode;                  // Код клиента
                    sheet.Cells[startAddPos, Global._clnClnNamePos] = item.ClientName;                  // Название клиента
                    sheet.Cells[startAddPos, Global._clnMngNamePos] = item.ManagerName;                 // Название менеджера
                    if (ordEmlData.OrdEmlDict.TryGetValue(item.ClientName, out emailValue))             // Если найдена запись с email клиента?
                    {
                        sheet.Cells[startAddPos, Global._clnEmailPos] = emailValue;                     // Определить email адреса клиента
                    }
                    sheet.Cells[startAddPos, Global._clnUserPos] = Environment.UserName;                // Имя пользователя
                    sheet.Cells[startAddPos, Global._clnAddTimePos] = DateTime.Now;                     // Дата и время добавления записи
                    UpdateAddedRowColor(sheet, startAddPos, Global._clnUpdTimePos);
                    startAddPos++;
                    addCount++;
                }
            }

            Global.OutputLine(string.Format("- Добавлено новых записей: {0}", addCount));
            Global.OutputLine(string.Format("- Обновлено записей: {0}", updCount));

            Global.OutputLine("");
            Global.OutputLine(string.Format("--> Дополнение данных по Статусам"));
            
            sheet = (Excel.Worksheet)book.Worksheets[Global._Statuses];             // Лист СТАТУСЫ

            // Поиск позиции для добавления первой новой строки на лист Статусы
            for (startAddPos = 2; sheet.Cells[startAddPos, 1].Value != null; startAddPos++) { }
            Global.OutputLine(string.Format("Номер стартовой строки для добавления новых строк в таблицу '{0}({1})': {2}", Path.GetFileName(Global.FileName_DebtorStatus), Global._Statuses, startAddPos));
            
            addCount = 0;

            foreach (DbtDblItem item in debtorData.DbtDtlDict.Values)               // Просмотр основной коллекции с данными по дебиторке
            {
                string key = item.ClientCode + item.UnitName;                       // КЛЮЧ поиска
                StsStsItem statusItem;
                if (!stsStsData.StsStsDict.TryGetValue(key, out statusItem))        // Если не найдена запись в базе статусов?
                {
                    sheet.Cells[startAddPos, Global._stsUntNamePos] = item.UnitName;                   // БЕ
                    sheet.Cells[startAddPos, Global._stsClnCodePos] = item.ClientCode;                 // Код клиента
                    sheet.Cells[startAddPos, Global._stsClnNamePos] = item.ClientName;                 // Название клиента
                    sheet.Cells[startAddPos, Global._stsUserPos] = Environment.UserName;               // Имя пользователя
                    sheet.Cells[startAddPos, Global._stsAddTimePos] = DateTime.Now;                    // Дата и время добавления записи
                    UpdateAddedRowColor(sheet, startAddPos, Global._stsUpdTimePos);
                    startAddPos++;
                    addCount++;
                }
            }
            
            Global.OutputLine(string.Format("- Добавлено новых записей: {0}", addCount));

            book.Save();
            book.Close();
        }

        // **********************************************************************
        // АКТУАЛИЗАЦИЯ СТАТУСОВ на основе данных Дебиторки
        public void ProcessDebtorData()
        {
            Global.OutputLine("");
            Global.OutputLine(string.Format("==>> АКТУАЛИЗАЦИЯ СТАТУСОВ и РАССЫЛКА ИЗВЕЩЕНИЙ"));

            Excel.Workbook book = null;
            Excel.Worksheet sheet = null;

            try
            {
                book = excelApp.Workbooks.Open(Global.FileName_DebtorStatus, Editable: true);       // Открытие файла со статусными данными
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("*** Ошибка: Не удалось открыть файл: '{0}'. Сообщение: '{1}'", Global.FileName_DebtorStatus, ex.Message));
                Global.IsFatalError = true;
                return;                     // -->>
            }

            sheet = (Excel.Worksheet)book.Worksheets[Global._Statuses];                             // Лист СТАТУСЫ

            StsClnItem clientItem;
            DbtDblItem debetorkaItem;
            
            TimeSpan? period;               // Период ожидания погашения просроченной задолженности

            Excel.Range cell;

            this.warningMsgCount = 0;       // Счетчик предупреждений о возможной приостановке отгрузок
            this.eraseMsgCount = 0;         // Счетчик сброса даты выдачи предупреждения
            this.stopMsgCount = 0;          // Счетчик извещений о приостановке отгрузок
            this.startMsgCount = 0;         // Счетчик извещений о возобновлении отгрузок

            this.currDate = DateTime.Today;                                                         // Текущая дата для формирования Документа с извещением

            Excel.Range rg = sheet.ListObjects.get_Item("Статусы_tb").DataBodyRange.Rows;
            rg.Interior.Color = Global.DefaultRowColor;                                            // Восстановить цвет фона строк таблицы
            long itemCount = 0;
           
            foreach (StsStsItem statusItem in stsStsData.StsStsDict.Values)                         // Просмотр дополненной таблицы статусов
            {
                itemCount++;
                Console.Write(string.Format("{0:###0}: {1}                                    \r", itemCount, statusItem.UnitName + " " + statusItem.ClientName));
                this.isUpdateRow = false;
                if (stsClnData.StsClnDict.TryGetValue(statusItem.ClientCode, out clientItem))       // Если по этому клиенту работаем (клиент не закрыт)?
                {
                    string key = statusItem.ClientCode + statusItem.UnitName;                       // КЛЮЧ поиска в Дебиторке: Код клиента + БЕ
                    if (debtorData.DbtDtlDict.TryGetValue(key, out debetorkaItem))                  // Если есть Дебиторка по этому ключу?
                    {
                        // Полная задолженность в Рублях
                        cell = sheet.Cells[statusItem.RowNum, Global._stsRubTotSumPos];             // Обновляемая ячейка
                        cell.Value = debetorkaItem.RubTotSum;                                       // Сумма полной задолженности в Рублях!
                        if (debetorkaItem.RubTotSum != statusItem.RubTotSum)                        // Были изменения RubTotSum?
                        {
                            UpdateCellColor(cell);
                        }

                        // Полная задолженность в EUR
                        cell = sheet.Cells[statusItem.RowNum, Global._stsEurTotSumPos];             // Обновляемая ячейка
                        if (debetorkaItem.EurTotSum != 0)
                        {
                            cell.Value = debetorkaItem.EurTotSum;                                   // Сумма полной задолженности в EUR!
                        }
                        else
                        {
                            cell.Value = null;
                        }
                        if (debetorkaItem.EurTotSum != statusItem.EurTotSum)                        // Были изменения EurTotSum?
                        {
                            UpdateCellColor(cell);
                        }

                        // Полная задолженность в USD
                        cell = sheet.Cells[statusItem.RowNum, Global._stsUsdTotSumPos];             // Обновляемая ячейка
                        if (debetorkaItem.UsdTotSum != 0)
                        {
                            cell.Value = debetorkaItem.UsdTotSum;                                   // Сумма полной задолженности в USD!
                        }
                        else
                        {
                            cell.Value = null;
                        }
                        if (debetorkaItem.UsdTotSum != statusItem.UsdTotSum)                        // Были изменения UsdTotSum?
                        {
                            UpdateCellColor(cell);
                        }

                        // Просроченная задолженность в Рублях
                        cell = sheet.Cells[statusItem.RowNum, Global._stsRubDueSumPos];             // Обновляемая ячейка
                        if (debetorkaItem.RubDueSum != 0)
                        {
                            cell.Value = debetorkaItem.RubDueSum;                                   // Сумма просроченной задолженности в Рублях!
                        }
                        else
                        {
                            cell.Value = null;
                        }
                        if (debetorkaItem.RubDueSum != statusItem.RubDueSum)                        // Были изменения RubDueSum?
                        {
                            UpdateCellColor(cell);
                        }

                        // Просроченная задолженность в EUR
                        cell = sheet.Cells[statusItem.RowNum, Global._stsEurDueSumPos];             // Обновляемая ячейка
                        if (debetorkaItem.EurDueSum != 0)
                        {
                            cell.Value = debetorkaItem.EurDueSum;                                   // Сумма просроченной задолженности в EUR!
                        }
                        else
                        {
                            cell.Value = null;
                        }
                        if (debetorkaItem.EurDueSum != statusItem.EurDueSum)                        // Были изменения EurDueSum?
                        {
                            UpdateCellColor(cell);
                        }

                        // Просроченная задолженность в USD
                        cell = sheet.Cells[statusItem.RowNum, Global._stsUsdDueSumPos];             // Обновляемая ячейка
                        if (debetorkaItem.UsdDueSum != 0)
                        {
                            cell.Value = debetorkaItem.UsdDueSum;                                   // Сумма просроченной задолженности в USD!
                        }
                        else
                        {
                            cell.Value = null;
                        }
                        if (debetorkaItem.UsdDueSum != statusItem.UsdDueSum)                        // Были изменения UsdDueSum?
                        {
                            UpdateCellColor(cell);
                        }

                        // Вычисление событий и статусов
                        if (statusItem.date1 == null)               // Нет даты выдачи Предупреждения (Дата1) ?
                        {
                            if (statusItem.date2 == null)           // Нет даты Приостановки отгрузок (Дата2) ?
                            {
                                if (debetorkaItem.RubDueSum > Global.MIN_SUM)       // Просроченная задолженность > MIN_SUM (выше порога чувствительности)?
                                {
                                    // MSG: Выдать сообщение с предупреждением о возможной приостановке отгрузок!
                                    cell = sheet.Cells[statusItem.RowNum, Global._stsData1Pos];         // Обновляемая ячейка Дата1
                                    cell.Value = this.currDate;                                         // Дата выдачи Предупреждения!
                                    UpdateTrafficColor(cell, Traffic.Yellow);
                                    SendWarningMsg(debetorkaItem, statusItem);      // ПРЕДУПРЕЖДЕНИЕ клиенту о возможной приостановке отгрузок
                                }
                            }
                            else                                    // Есть дата Приостановки отгрузок (Дата2)
                            {
                                // Должны быть обе даты (Дата1 и Дата2)
                                Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})'. Есть Дата2 и нет Дата1 в строке {2}", Global.FileName_DebtorStatus, Global._Statuses, statusItem.RowNum));
                                Global.IsNoncriticzlError = true;
                            }
                        }
                        else                                        // Есть дата выдачи Предупреждения (Дата1) !
                        {
                            if (statusItem.date2 == null)           // Нет даты приостановки отгрузок (Дата2) ?
                            {
                                period = this.currDate - statusItem.date1;                              // Период погашения
                                if (period != null && period.Value.Days >= Global.DEAD_LINE)            // Период погашения истек ?
                                {
                                    if (debetorkaItem.RubDueSum > Global.MIN_SUM)      // Просроченная задолженность выше порога чувствительности ?
                                    {
                                        // MSG: Выдать сообщение о ПРИОСТАНОВКЕ ОТГРУЗОК!
                                        cell = sheet.Cells[statusItem.RowNum, Global._stsData2Pos];     // Обновляемая ячейка Дата2
                                        cell.Value = this.currDate;                                     // Дата Приостановки отгрузок
                                        UpdateTrafficColor(cell, Traffic.Red);
                                        SendStopMsg(statusItem);    // Извещение о ПРИОСТАНОВКЕ отгрузок
                                    }
                                    else
                                    {
                                        // Сброс даты выдачи предупреждения! (СБРОС СТАТУСА)
                                        cell = sheet.Cells[statusItem.RowNum, Global._stsData1Pos];
                                        cell.Value = null;          // Сделать Пусто Дата1 !
                                        UpdateTrafficColor(cell, Traffic.Green);
                                        this.eraseMsgCount++;
                                    }
                                }
                            }
                            else                                    // Есть дата приостановки отгрузок!
                            {
                                if (debetorkaItem.RubDueSum <= Global.MIN_SUM)     // Нет просроченной задолженности ?
                                {
                                    // СБРОС СТАТУСОВ
                                    cell = sheet.Cells[statusItem.RowNum, Global._stsData1Pos];
                                    cell.Value = null;              // Сделать Пусто Дата1 !
                                    UpdateTrafficColor(cell, Traffic.Green);
                                    cell = sheet.Cells[statusItem.RowNum, Global._stsData2Pos];
                                    cell.Value = null;              // Сделать Пусто Дата2 !
                                    UpdateTrafficColor(cell, Traffic.Green);
                                    SendStartMsg(statusItem);       // Извещение о ВОЗОБНОВЛЕНИИ отгрузок
                                }
                            }
                        }
                    }
                    else            // Нет дебиторки!
                    {
                        // Очистить полную задолженность
                        if (statusItem.RubTotSum != 0)
                        {
                            cell = sheet.Cells[statusItem.RowNum, Global._stsRubTotSumPos];
                            cell.Value = null;                      // Сделать Пусто Рубли !
                            UpdateCellColor(cell);
                        }
                        if (statusItem.EurTotSum != 0)
                        {
                            cell = sheet.Cells[statusItem.RowNum, Global._stsEurTotSumPos];
                            cell.Value = null;                      // Сделать Пусто EUR !
                            UpdateCellColor(cell);
                        }
                        if (statusItem.EurTotSum != 0)
                        {
                            cell = sheet.Cells[statusItem.RowNum, Global._stsUsdTotSumPos];
                            cell.Value = null;                      // Сделать Пусто USD !
                            UpdateCellColor(cell);
                        }

                        // Очистить просроченную задолженность
                        if (statusItem.RubDueSum != 0)
                        {
                            cell = sheet.Cells[statusItem.RowNum, Global._stsRubDueSumPos];
                            cell.Value = null;                      // Сделать Пусто Рубли !
                            UpdateCellColor(cell);
                        }
                        if (statusItem.EurDueSum != 0)
                        {
                            cell = sheet.Cells[statusItem.RowNum, Global._stsEurDueSumPos];
                            cell.Value = null;                      // Сделать Пусто EUR !
                            UpdateCellColor(cell);
                        }
                        if (statusItem.UsdDueSum != 0)
                        {
                            cell = sheet.Cells[statusItem.RowNum, Global._stsUsdDueSumPos];
                            cell.Value = null;                      // Сделать Пусто USD !
                            UpdateCellColor(cell);
                        }

                        if (statusItem.date1 == null)               // Нет даты выдачи предупреждения?
                        {
                            if (statusItem.date2 == null)           // Нет даты приостановки отгрузок?
                            {
                                // Ничего не делать!
                            }
                            else
                            {
                                Global.OutputLine(string.Format("### ошибка! В файле '{0}({1})'. Есть Дата2 и нет Дата1 в строке {2}", Global.FileName_DebtorStatus, Global._Statuses, statusItem.RowNum));
                                Global.IsNoncriticzlError = true;
                            }
                        }
                        else                                        // Есть дата выдачи предупреждения!
                        {
                            if (statusItem.date2 == null)           // Нет даты приостановки отгрузок?
                            {
                                // Сброс даты выдачи предупреждения! (СБРОС СТАТУСА)
                                cell = sheet.Cells[statusItem.RowNum, Global._stsData1Pos];
                                cell.Value = null;                  // Сделать Пусто Дата1 !
                                UpdateTrafficColor(cell, Traffic.Green);
                                this.eraseMsgCount++;
                            }
                            else                                    // Есть дата ПРИОСТАНОВКЕ отгрузок!
                            {
                                // СБРОС СТАТУСОВ
                                cell = sheet.Cells[statusItem.RowNum, Global._stsData1Pos];
                                cell.Value = null;                  // Сделать Пусто Дата1 !
                                UpdateTrafficColor(cell, Traffic.Green);
                                cell = sheet.Cells[statusItem.RowNum, Global._stsData2Pos];
                                cell.Value = null;                  // Сделать Пусто Дата2 !
                                UpdateTrafficColor(cell, Traffic.Green);
                                SendStartMsg(statusItem);           // Извещение о ВОЗОБНОВЛЕНИИ отгрузок
                            }
                        }
                    }
                }
                else        // По этому клиенту не работаем (закрыт!)
                {
                    UpdateSkipRowColor(sheet, statusItem.RowNum, Global._stsUpdTimePos);      // Подкрасить выключенную строку
                }

                if (this.isUpdateRow)
                {
                    SetUpdateMarker(sheet, statusItem.RowNum);
                }
            }

            Global.OutputLine(string.Format("Выдано предупреждений о возможной приостановке отгрузок: {0}", this.warningMsgCount));
            Global.OutputLine(string.Format("Количество сброшенных дат о ранее выданном предупреждении: {0}", this.eraseMsgCount));
            Global.OutputLine(string.Format("Выдано извещений о приостановке отгрузок: {0}", this.stopMsgCount));
            Global.OutputLine(string.Format("Выдано извещений о возобновлении отгрузок: {0}", this.startMsgCount));

            book.Save();        // Сохранить изменения в excel файле
            book.Close();       // Закрыть файл
        }

        // Формирование полного имени файла с извещением 1 (предупреждение о возможной приостановке отгрузок)
        private string GetNotice1FullFileName(string folderName, string clientName, string unitName)
        {
            string dateStr = string.Format("{0:yyyy.MM.dd HH:mm:ss}", DateTime.Now);
            dateStr = dateStr.Replace(':', '-');
            return Path.Combine(Global.FolderName_Notice, "Извещение1 для " + clientName + "(" + unitName + ") " + dateStr + ".pdf");
        }

        private void MakeWarningMsg(DbtDblItem dbtItem, StsStsItem stsItem, StsClnItem clnItem, string managerEmail, out string noticeFileName)
        {
            string templateFileName;
            string documentFileName;
            string stampFileName;
            UnitItem unitItem;
            noticeFileName = null;

            if (this.unitData.UnitDict.TryGetValue(stsItem.UnitName, out unitItem))
            {
                templateFileName = Path.Combine(Global.FolderName_Template, Global.FileName_TemplateNotice1);                   // Имя файла с Шаблоном извещения 1
                documentFileName = GetNotice1FullFileName(Global.FolderName_Notice, clnItem.ClientName, unitItem.UnitName);     // Имя файла с документом Извещения 1
                stampFileName = Path.Combine(Global.FolderName_Template, unitItem.StampFileName);                               // Имя файла с изображением печати и подписи для БЕ

                WordDocument oDoc = new WordDocument(this.wordApp, templateFileName);
                oDoc.ReplaceString("$$$unit", unitItem.UnitTitle);      // Замена в шаблоне ключа на значение
                oDoc.ReplaceString("$$$client", clnItem.ClientName);
                oDoc.ReplaceString("$$$innkpp", unitItem.InnCode);
                oDoc.ReplaceString("$$$account", unitItem.AccCode);
                oDoc.ReplaceString("$$$bikks", unitItem.BikCode);
                oDoc.ReplaceString("$$$mail", unitItem.UnitEmail);
                oDoc.ReplaceString("$$$sum", string.Format("{0:# ### ##0.00}", dbtItem.RubDueSum));
                oDoc.ReplaceString("$$$date2", string.Format("{0:dd.MM.yyyy}", this.currDate.AddDays(Global.DEAD_LINE)));
                oDoc.ReplaceString("$$$date1", string.Format("{0:d MMMM yyyy}", this.currDate));
                oDoc.ReplaceString("$$$manager", unitItem.UnitManager);

                oDoc.SaveAndClose(documentFileName, stampFileName);     // Сохранение и закрытие созданного документа
                noticeFileName = documentFileName;                      // Имя файла с созданным Извещением 1    
                oDoc = null;
            }
        }

        // Рассылка предупреждения клиенту о возможной приостановке отгрузок
        private void SendWarningMsg(DbtDblItem dbtItem, StsStsItem stsItem)
        {
            StsClnItem clnItem;
            string managerEmail;
            string noticeFileName;          // Имя созданного файла с Извещением 1 для рассылки по email
            
            if (this.stsClnData.StsClnDict.TryGetValue(stsItem.ClientCode, out clnItem))
            {
                if (!string.IsNullOrWhiteSpace(clnItem.EmailValue))                 // Если есть email клиента ?
                {
                    if (clnItem.ManagerName != null && this.managerData.ManagerDict.TryGetValue(clnItem.ManagerName, out managerEmail))
                    {
                        Global.OutputLine(string.Format("@@@ MSG! Предупреждение клиенту: '{0}({1})', email клиента: '{2}', email менеджера: '{3}'", stsItem.ClientName, stsItem.UnitName, clnItem.EmailValue, managerEmail));
                        MakeWarningMsg(dbtItem, stsItem, clnItem, managerEmail, out noticeFileName);        // Формирование текста письма с Извещением 1
                        this.warningMsgCount++;
                        // , Tatyana.Karaseva@sibelco.com
                        //Global.mailer.SendNotice("alexandr.pyatkov@sibelco.com", noticeFileName, stsItem.UnitName, stsItem.ClientName);
                    }
                    else
                    {
                        Global.OutputLine(string.Format("@@@ MSG? Предупреждение клиенту: '{0}({1})', email клиента: '{2}'", stsItem.ClientName, stsItem.UnitName, clnItem.EmailValue));
                        this.warningMsgCount++;
                        Global.OutputLine(string.Format("### ошибка! Менеджер без email '{0}' не получил копию предупреждения.", clnItem.ManagerName));
                        Global.IsNoncriticzlError = true;
                    }
                }
                else
                {
                    Global.OutputLine(string.Format("### ошибка! Не отослано предупреждение клиенту '{0}({1})'. Отсутствует email!",stsItem.ClientName, stsItem.UnitName));
                    Global.IsNoncriticzlError = true;
                }
            }
        }

        // Рассылка извещения клиенту о приостановке отгрузок
        private void SendStopMsg(StsStsItem stsItem)
        {
            StsClnItem clnItem;
            string managerEmail;

            if (this.stsClnData.StsClnDict.TryGetValue(stsItem.ClientCode, out clnItem))
            {
                if (!string.IsNullOrWhiteSpace(clnItem.EmailValue))         // Если есть email клиента ?
                {
                    if (clnItem.ManagerName != null && this.managerData.ManagerDict.TryGetValue(clnItem.ManagerName, out managerEmail))
                    {
                        Global.OutputLine(string.Format("@@@ MSG! Извещение клиенту о приостановке отгрузок: '{0}({1})', email клиента: '{2}', email менеджера: '{3}'", stsItem.ClientName, stsItem.UnitName, clnItem.EmailValue, managerEmail));
                        this.stopMsgCount++;
                    }
                    else
                    {
                        Global.OutputLine(string.Format("@@@ MSG! Извещение о приостановке отгрузок клиенту : '{0}({1})', email клиента: '{2}'", stsItem.ClientName, stsItem.UnitName, clnItem.EmailValue));
                        this.stopMsgCount++;
                        Global.OutputLine(string.Format("### ошибка! Менеджер без email '{0}' не получил копию извещения!", clnItem.ManagerName));
                        Global.IsNoncriticzlError = true;
                    }
                }
                else
                {
                    Global.OutputLine(string.Format("### ошибка! Не отослано извещение о приостановке отгрузок клиенту: '{0}({1})'. Отсутствует email.", stsItem.ClientName, stsItem.UnitName));
                    Global.IsNoncriticzlError = true;
                }
            }
        }

        // Рассылка извещения клиенту о возобновлении отгрузок
        private void SendStartMsg(StsStsItem stsItem)
        {
            StsClnItem clnItem;
            string managerEmail;

            if (this.stsClnData.StsClnDict.TryGetValue(stsItem.ClientCode, out clnItem))
            {
                if (!string.IsNullOrWhiteSpace(clnItem.EmailValue))         // Если есть email клиента ?
                {
                    if (clnItem.ManagerName != null && this.managerData.ManagerDict.TryGetValue(clnItem.ManagerName, out managerEmail))
                    {
                        Global.OutputLine(string.Format("@@@ MSG! Извещение клиенту о возобновлении отгрузок: '{0}({1})', email клиента: '{2}', email менеджера: '{3}'", stsItem.ClientName, stsItem.UnitName, clnItem.EmailValue, managerEmail));
                        this.startMsgCount++;
                    }
                    else
                    {
                        Global.OutputLine(string.Format("@@@ MSG! Извещение о возобновлении отгрузок клиенту : '{0}({1})', email клиента: '{2}'", stsItem.ClientName, stsItem.UnitName, clnItem.EmailValue));
                        this.startMsgCount++;
                        Global.OutputLine(string.Format("### ошибка! Менеджер без email '{0}' не получил копию извещения!", clnItem.ManagerName));
                        Global.IsNoncriticzlError = true;
                    }
                }
                else
                {
                    Global.OutputLine(string.Format("### ошибка! Не отослано извещение о возобновлении отгрузок клиенту: '{0}({1})'. Отсутствует email.", stsItem.ClientName, stsItem.UnitName));
                    Global.IsNoncriticzlError = true;
                }
            }
        }

        // Существует ли значение в строке
        private bool Exist(string item)
        {
            return !string.IsNullOrWhiteSpace(item);
        }

        // Обновить значение в ячейке
        private void UpdateCellValue(Excel.Worksheet sheet, long rowPos, int colPos, string newValue)
        {
            Excel.Range cell = sheet.Cells[rowPos, colPos];                     // Обновляемая ячейка
            cell.Value = newValue;                                              // Новое значение ячейки
            cell.Interior.Color = Global.UpdatedCellColor;                      // Цвет для подкраски обновленной ячейки
            sheet.Cells[rowPos, Global._clnUserPos] = Environment.UserName;     // Пользователь
            sheet.Cells[rowPos, Global._clnUpdTimePos] = DateTime.Now;          // Дата и время обновления
        }

        // Обновить цвет фона добавленной (новой) строки
        private void UpdateAddedRowColor(Excel.Worksheet sheet, long rowPos, int lastColPos)
        {
            Excel.Range cell1 = sheet.Cells[rowPos, 1];
            Excel.Range cell2 = sheet.Cells[rowPos, lastColPos];
            Excel.Range rg = sheet.get_Range(cell1, cell2);
            rg.Interior.Color = Global.AddedRowColor;                           // Цвет для подкраскидобавленных (новых) строк
        }

        // Обновить цвет фона пропущенной строки
        private void UpdateSkipRowColor(Excel.Worksheet sheet, long rowPos, int lastColPos)
        {
            Excel.Range cell1 = sheet.Cells[rowPos, 1];
            Excel.Range cell2 = sheet.Cells[rowPos, lastColPos];
            Excel.Range rg = sheet.get_Range(cell1, cell2);
            rg.Interior.Color = Global.SkipRowColor;                          // Цвет для подкраски пропущенных строк
        }

        // Отметить обновление
        private void SetUpdateMarker(Excel.Worksheet sheet, long rowNum)
        {
            sheet.Cells[rowNum, Global._stsUserPos] = Environment.UserName;     // Пользователь
            sheet.Cells[rowNum, Global._stsUpdTimePos] = DateTime.Now;          // Дата и время обновления
            UpdateCellColor(sheet.Cells[rowNum, Global._stsUpdTimePos]);
        }

        private void UpdateCellColor(Excel.Range cell)
        {
            cell.Interior.Color = Global.UpdatedCellColor;                      // Цвет для подкраски обновленных ячеек с данными
            this.isUpdateRow = true;
        }

        // Обновление цвета фона ячеек с датами по принципу светофора
        private void UpdateTrafficColor(Excel.Range cell, Traffic traffic)
        {   
            switch (traffic)
            {
                case Traffic.Green:
                    cell.Interior.Color = Global.TrafficGreenColor;
                    break;
                case Traffic.Yellow:
                    cell.Interior.Color = Global.TrafficYellowColor;
                    break;
                case Traffic.Red:
                    cell.Interior.Color = Global.TrafficRedColor;
                    break;
            }
            this.isUpdateRow = true;
        }
    }
}
