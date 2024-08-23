using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
using WordRange = Microsoft.Office.Interop.Word.Range;


namespace WordLib
{
    public class WordApp
    {
        private Application _app { get; set; }
        private Document _doc { get; set; }
        private WordRange _range { get; set; }
        private object oMissing = System.Reflection.Missing.Value;
        private object oEndOfDoc = "\\endofdoc";

        public WordApp(bool visible = false)
        {
            _app = new Application
            {
                Visible = visible
            };
        }

        public void OpenDoc(string file = null)
        {
            if (file == null)
            {
                _doc = _app.Documents.Add();
            }
            else
            {
                _doc = _app.Documents.Open(file);
            }
        }

        public void CreateDoc()
        {
            _doc = _app.Documents.Add();
        }

        /// <summary>
        /// возвращает параграфы в виде списка строк
        /// </summary>
        /// <returns></returns>
        public List<string> ReadContent()
        {
            var list = new List<string>();
            foreach (Paragraph paragraph in _doc.Paragraphs)
            {
                list.Add(paragraph.Range.Text);
            }
            return list;
        }

        /// <summary>
        /// форматирование для форм иц
        /// </summary>
        public void FormattDoc()
        {
            GetAllContent();
            if (_range != null)
            {
                _range.Font.Size = 8;
                _range.Font.Scaling = 80;
                _range.Font.Name = "Courier New";
                _doc.PageSetup.LeftMargin = 28;
                _doc.PageSetup.RightMargin = 28;
                _doc.PageSetup.TopMargin = 28;
                _doc.PageSetup.BottomMargin = 28;
                _doc.Content.ParagraphFormat.LineSpacing = 12;
                _doc.Content.ParagraphFormat.SpaceAfter = 0;
            }
        }

        /// <summary>
        /// добавить содержимое между указанными позициями
        /// </summary>
        /// <param name="content"></param>
        /// <param name="startIndex"></param>
        /// <param name="endIndex"></param>
        public void AddContent(string content, int startIndex, int endIndex)
        {
            WordRange rng = _doc.Range(startIndex, endIndex);
            rng.Text = content;
        }

        /// <summary>
        /// добавить параграф в начало документа
        /// </summary>
        /// <param name="content"></param>
        public void AddContentToBegin(string content)
        {
            Paragraph para = _doc.Content.Paragraphs.Add();
            para.Range.Text = content;
        }

        /// <summary>
        /// добавить параграф в конец документа
        /// </summary>
        /// <param name="content"></param>
        public void AddContentToEnd(string content)
        {
            Paragraph para = _doc.Content.Paragraphs.Add(oEndOfDoc);
            para.Range.Text = content;
        }

        /// <summary>
        /// найти и заменить
        /// </summary>
        /// <param name="findText">что найти</param>
        /// <param name="replaceText">чем заменить</param>
        public void SearchReplace(string findText, string replaceText)
        {
            Find findObject = _app.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = findText;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replaceText;

            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing,
                ref replaceAll, oMissing, oMissing, oMissing, oMissing);
        }

        /// <summary>
        /// ищет текст между двумя указанными строками (before и after) в документе Word,
        /// а затем возвращает найденный текст
        /// </summary>
        /// <param name="before">текс до искомого фрагмента</param>
        /// <param name="after">текст после искомого фрагмента</param>
        /// <returns></returns>
        public object CopyBetween(string before, string after)
        {
            object result = null;
            var range = _doc.Content;
            var start = 0;
            var end = range.End;
            range = _doc.Range(start, end);
            if (range.Find.Execute(before))
            {
                start = range.End;
                end = start;
                range = _doc.Range(start, end);
                range.Select();
                if (range.Find.Execute(after))
                {
                    end = range.Start;
                    if (end > start)
                    {
                        result = _doc.Range(start, end).Text.Trim();
                        return result;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// выделяет строки, содержащие указанные текст
        /// </summary>
        /// <param name="text">искомый текст</param>
        public void HighLightRowsWithText(string text)
        {
            var ranges = GetRowsWithText(text);
            foreach(var range in ranges)
            {
                range.Font.Bold = 1;
                range.Font.Underline = WdUnderline.wdUnderlineSingle;
            }
        }

        public void SaveDoc()
        {
            _doc.Save();
        }

        public void SaveDocAs(string file)
        {
            _doc.SaveAs(file);
        }

        public void OpenDocInUTF8(string file)
        {
            if (!File.Exists(file))
            {
                throw new Exception();
            }
            // Encoding.GetEncoding(1251).CodePage
            _doc = _app.Documents.Open(
                file, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, Encoding.UTF8.CodePage, oMissing,
                oMissing, oMissing, oMissing, oMissing);
        }

        public void SaveDocAsCyrillicEncoding(string file)
        {
            // Encoding.UTF8.CodePage
            _doc.SaveAs(
                file, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, Encoding.GetEncoding(1251).CodePage,
                oMissing, oMissing, oMissing, oMissing);
        }

        public void SaveAsPdf(string file)
        {
            _doc.ExportAsFixedFormat(file.ToString(),
                    WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                    WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true,
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, oMissing);
        }

        public void CloseDoc()
        {
            if (_doc != null)
            {
                _doc.Close();
                Marshal.ReleaseComObject(_doc);
                _doc = null;
            }
        }

        public void Quit()
        {
            _app.Quit();
            if (_app != null)
            {
                Marshal.ReleaseComObject(_app);
                _app = null;
            }
        }

        private List<WordRange> GetRowsWithText(string text)
        {
            List<WordRange> result = new List<WordRange>();
            var range = _doc.Content;
            var start = 0;
            var end = range.End;
            range = _doc.Range(start, end);
            while (range.Find.Execute(text))
            {
                start = range.Start;
                end = start;
                range = _doc.Range(start, end);
                range.Select();
                if (range.Find.Execute(@"^p"))
                {
                    end = range.Start;
                    if (end > start)
                    {
                        result.Add(_doc.Range(start, end));
                        range = _doc.Range(end, _doc.Content.End);
                    }
                }
            }
            return result;
        }




        private void GetAllContent()
        {
            var start = _doc.Content.Start;
            var end = _doc.Content.End;
            _range = _doc.Range(start, end);
        }
    }
}
