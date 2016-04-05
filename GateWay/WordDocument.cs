using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace GateWay
{
    // Word запускается ОТДЕЛЬНЫМ ПРОЦЕССОМ, класс просто управляет им через библиотеку Word Interoperability, 
    // на компьютере должен быть установлен офис, в проекте должна быть ссылка на Microsoft.Office.Interop.Word, 
    // соотвествующая библиотека .dll должна быть в папке с программой
    // документ MS WORD, позволяет создать новый документ по шаблону, произвести поиск и замену строк (одно вхождение или все), 
    // изменить видимость документа, закрыть документ
    public class WordDocument
    {
        // фиксированные параметры для передачи приложению Word
        private Object wordMissing = System.Reflection.Missing.Value;
        // если использовать Word.Application и Word.Document получим предупреждение от компиллятора
        private Word._Application wordApp;              // Приложение Word (передается через конструктор)
        private Word._Document wordDocument;            // Документ, созданный по шаблону
        private Object templatePathObj;                 // Объект с шаблоном создаваемого документа

        // конструктор, создаем по шаблону, потом возможно расширение другими вариантами
        public WordDocument(Word._Application wordApp, string templatePath)
        {
            this.wordApp = wordApp;                     // Приложение Word
            this.templatePathObj = templatePath;        // Путь к файлу с шаблоном
            try
            {
                this.wordDocument = this.wordApp.Documents.Add(ref templatePathObj, ref wordMissing, ref wordMissing, ref wordMissing);  // Создание документа по шаблону
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("*** Ошибка: Не удалось открыть файл с шаблоном документа: '{0}'. Сообщение: '{1}'", (string)templatePath, ex.Message));
                Global.IsFatalError = true;
            }
        }

        // ПОИСК И ЗАМЕНА ЗАДАННОЙ СТРОКИ
        public void ReplaceString(string strToFind, string replaceStr)
        {
            object strToFindObj = strToFind;
            object replaceStrObj = replaceStr;
            object replaceTypeObj;
            Word.Range wordRange;

            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            try
            {
                // обходим все разделы документа
                for (int i = 1; i <= wordDocument.Sections.Count; i++)
                {
                    // берем всю секцию диапазоном
                    wordRange = wordDocument.Sections[i].Range;
                    // выполняем метод поискаи  замены обьекта диапазона ворд
                    wordRange.Find.Execute(ref strToFindObj, ref wordMissing, ref wordMissing, ref wordMissing,
                                        ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref replaceStrObj,
                                        ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                }
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("*** Ошибка при  выполнении поиска и замены в документе Word: '{0}'. Сообщение: '{1}'", wordDocument.Name, ex.Message));
                Global.IsFatalError = true;
            }
        }
        
        // Сохранить файл с извещением в формате PDF
        public void SaveAndClose(string documentFileName, string stampFileName)
        {
            var shape = this.wordDocument.Bookmarks["ПЕЧАТЬ"].Range.InlineShapes.AddPicture(stampFileName, false, true);
            shape.Width = 120;
            shape.Height = 120;

            Object documentPathObj;
            try
            {
                documentPathObj = documentFileName;        // Путь к файлу с документом
                this.wordDocument.Activate();
                this.wordDocument.SaveAs2(ref documentPathObj, Word.WdSaveFormat.wdFormatPDF);
                this.wordDocument.Close(false);
                this.wordDocument = null;
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("*** Ошибка при попытке сохранить и закрыть документ: '{0}'", documentFileName));
                Global.OutputLine(string.Format("*** Сообщение об ошибке: '{0}'", ex.Message));
                Global.IsFatalError = true;
            }
        }
    }
}
