using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GateWay
{
    public static class ExcelImport
    {
        public static DataTable ImportExcelXLS(string fileName, string tableName, bool hasHeaders = true)
        {
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn;
            string _tableName = tableName + "$";
            bool IsNormalExtention = false;

            DataTable resultTable = null;

            if (!(File.Exists(fileName)))
            {
                Global.IsFatalError = true;
                Global.OutputLine(string.Format("*** Ошибка! Не найден файл с именем '{0}'", fileName));
            }
            else
            {
                string fileExtention = fileName.Substring(fileName.LastIndexOf('.')).ToLower();     // Расширение имени файла
                if (fileExtention == ".xlsx")
                {
                    strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=Read;Extended Properties=\"Excel 12.0 Xml;HDR={1};IMEX=0;ReadOnly=true;\"", fileName, HDR);
                    IsNormalExtention = true;
                }
                else if (fileExtention == ".xlsm")      // Если файл с макросами ?
                {
                    strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=Read;Extended Properties=\"Excel 12.0 Macro;HDR={1};IMEX=0;ReadOnly=true;\"", fileName, HDR);
                    IsNormalExtention = true;
                }
                else
                {
                    strConn = "";
                    Global.IsFatalError = true;
                    Global.OutputLine(string.Format("*** Ошибка! Файл '{0}' именет неправильное расширение '{1}'", Path.GetFileName(@fileName), fileExtention));
                }
                if (IsNormalExtention)                  // Если правильное расширение имени файла ?
                {
                    using (var conn = new OleDbConnection(strConn))
                    {
                        try
                        {
                            conn.Open();
                            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                            foreach (DataRow schemaRow in schemaTable.Rows)
                            {
                                string sheet = schemaRow["TABLE_NAME"].ToString();
                                sheet = sheet.Replace("'", "");     // Поправка (19.10.2015)
                                if (_tableName == sheet)    // Обрабатывать только лист с заданным именем TableName
                                {
                                    OleDbCommand cmd = new OleDbCommand("select * from [" + sheet + "]", conn);
                                    cmd.CommandType = CommandType.Text;
                                    resultTable = new DataTable(sheet);
                                    using (var da = new OleDbDataAdapter(cmd))
                                    {
                                        da.Fill(resultTable);
                                    }
                                    break;
                                }
                            }
                            if (resultTable == null)
                            {
                                Global.IsFatalError = true;
                                Global.OutputLine(string.Format("*** Ошибка! В файле '{0}' не найден лист с именем '{1}'", Path.GetFileName(fileName), tableName));
                            }
                            else
                            {
                                Global.OutputLine("");
                                Global.OutputLine(string.Format("--> Загрузка данных с листа '{0}' из excel файла '{1}'", tableName, Path.GetFullPath(@fileName)));
                            }
                        }
                        catch (Exception Ex)
                        {
                            Global.IsFatalError = true;
                            Global.OutputLine(string.Format("*** Ошибка! Файл: '{0}', Таблица: '{1}', Сообщение: '{2}'", Path.GetFileName(@fileName), tableName, Ex.Message));
                            resultTable = null;
                        }
                    }
                }
            }
            return resultTable;
        }
    }
}
