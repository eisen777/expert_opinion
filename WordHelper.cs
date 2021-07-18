using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace expert_opinion
{
    class WordHelper
    {
        //Получаем текущий рабочий каталог приложения, путь к текущей рабочей папке без обратной косой черты
        //(\) в конце.
        static string pathOfDir = Directory.GetCurrentDirectory();
        //указываем пути к шаблонам СИ
        object prowirl_200 = Path.Combine(pathOfDir, @"flowmeters\prowirl_200.docx");
        object USMGT400 = Path.Combine(pathOfDir, @"flowmeters\USM-GT-400.docx");
        object flowsick600xt = Path.Combine(pathOfDir, @"flowmeters\flowsick600-xt.docx");
        //указываем путь к папке с готовыми документами
        string target = @"c:\temp";
        private FileInfo _fileInfo;

        public WordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File no exists");
            }
        }
        //метод, производящий замену ключевых слов значениями и абзацами из ворда
        internal bool Process(Dictionary<string, string> items)
        {   
            //объявляем ссылку на шаблон МЭ
            Word.Application app = null;
            //заглушка для опциональных аргументов
            object oMissing = System.Reflection.Missing.Value;
            try
            {
                //создаем экземпляр класса приложения ворд для шаблона
                app = new Word.Application();
                
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                //создаем документ ворд для шаблона
                Word.Document doc = app.Documents.Open(file);
                //вставляем название работы и название документа в колонтитулы
                foreach (var item in items)
                {

                    Word.Find find = doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find;

                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);

                }
                //вставляем значения и абзацы (для средств измерений из других документов ворд) в основной текст
                foreach (var item in items)
                {
                    if(item.Key == "<flow_meter>" && item.Value == "Prowirl 200")
                    {
                        Word.Application wordFlow_meter = new Word.Application();
                        Word.Document sdoc = wordFlow_meter.Documents.Add(ref prowirl_200, ref oMissing, ref oMissing, ref oMissing);
                        sdoc.Range(ref oMissing, ref oMissing).Copy();
                        //я хз как вычислить номер абзаца, поэтому задаю жестко, тем более это константа
                        doc.Paragraphs[21].Range.Paste();
                    }
                    else if (item.Key == "<flow_meter>" && item.Value == "FLOWSIC600-XT")
                    {
                        Word.Application wordFlow_meter = new Word.Application();
                        Word.Document sdoc = wordFlow_meter.Documents.Add(ref flowsick600xt, ref oMissing, ref oMissing, ref oMissing);
                        sdoc.Range(ref oMissing, ref oMissing).Copy();
                        //я хз как вычислить номер абзаца, поэтому задаю жестко, тем более это константа
                        doc.Paragraphs[21].Range.Paste();
                    }
                    else if (item.Key == "<flow_meter>" && item.Value == "USM-GT-400")
                    {
                        Word.Application wordFlow_meter = new Word.Application();
                        Word.Document sdoc = wordFlow_meter.Documents.Add(ref USMGT400, ref oMissing, ref oMissing, ref oMissing);
                        sdoc.Range(ref oMissing, ref oMissing).Copy();
                        //я хз как вычислить номер абзаца, поэтому задаю жестко, тем более это константа
                        doc.Paragraphs[21].Range.Paste();
                    }
                    else
                    {
                        Word.Find find = app.Selection.Find;


                        find.Text = item.Key;
                        find.Replacement.Text = item.Value;



                        Object wrap = Word.WdFindWrap.wdFindContinue;
                        Object replace = Word.WdReplace.wdReplaceAll;

                        find.Execute(FindText: Type.Missing,
                            MatchCase: false,
                            MatchWildcards: false,
                            MatchSoundsLike: missing,
                            MatchAllWordForms: false,
                            Forward: true,
                            Wrap: wrap,
                            Format: false,
                            ReplaceWith: missing, Replace: replace);
                    }
                }
                //если каталога с готовыми документами не существует он будет создан
                if (!Directory.Exists(target))
                {
                    Directory.CreateDirectory(target);
                }
                Object newFileName = Path.Combine(target, DateTime.Now.ToString("yyMMdd") + " МЭ " + items["<title>"]);
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();


                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }
            return false;
        }

        
    }
}
