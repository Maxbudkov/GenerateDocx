using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using MySql.Data.MySqlClient;

namespace phpESH
{
    class Program
    {

        static Word.Application app = null;
        static Word.Document doc = null;

        static void Main(string[] args)
        {
            string config = @"config.cfg";
            List<String> fullpath = ReturnPath(config);
            List<String> bookmarks = ReturnBookmarks(config);
            List<String> replaceBookmarks = onDB();//ReturnReplacement(config);
            Word.Range temp = null;

            for (int i = 0; i < fullpath.Count; i++)
            {
                temp = Open(fullpath[i].ToString());
                Booking(temp, bookmarks, replaceBookmarks);
                Console.WriteLine("Сохраняем");
                doc.SaveAs2((Path.GetDirectoryName(fullpath[i].ToString()) + @"\ESH_" + Path.GetFileName(fullpath[i].ToString())));
                Console.WriteLine("Сохранил");
                Close();
            }


            Console.WriteLine("Done. ");
            Console.ReadKey();
        }

        private static List<String> ReturnPath(string config)
        {
            string[] fs = null;
            List<string> fullPath = new List<string>();
            try
            {
                fs = File.ReadAllLines(Path.GetFullPath(config));
                //Перебираем конфиг
                for (int i = 0; i < fs.Length; i++)
                {
                    //Встретили пути файлов
                    if (fs[i] == "[FullPath]")
                    {
                        for (int j = i + 1; j < fs.Length; j++)
                        {
                            if (fs.Length > (j))
                            {
                                if (fs[j][0] != '/')
                                {
                                    fullPath.Add(Path.GetFullPath(fs[j]).ToString());
                                    //Console.WriteLine("FP: " + fullPath[fullPath.Count - 1].ToString());
                                }
                                else break;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Ошибка чтения файла конфигурации. " +
                    "Проверьте файл config.cfg. Он должен находиться в папке с программой.");
                Console.WriteLine("Ошибка: " + e.Message);
            }
            return fullPath;
        }

        private static List<String> ReturnBookmarks(string config)
        {
            string temp = null;
            string[] fs = null;
            List<string> bookmarks = new List<string>();
            try
            {
                fs = File.ReadAllLines(Path.GetFullPath(config));
                //Перебираем конфиг построчно
                for (int i = 0; i < fs.Length; i++)
                {
                    //Встретили закладки
                    if (fs[i] == "[Bookmarks]")
                    {
                        //Начиная с раздела закладки
                        for (int j = i; j < fs.Length; j++)
                        {
                            //Если не в конце файла
                            if ((j + 1) < fs.Length)
                            {
                                //Если не встретили следующий раздел конфига
                                if (fs[j + 1][0] != '/')
                                {
                                    //Перебираем строку
                                    for (int k = 0; k < fs[j + 1].Length; k++)
                                    {
                                        if (fs[j + 1][k] == '=')
                                        {
                                            bookmarks.Add((temp).ToString());
                                            temp = null;
                                            break;
                                        }
                                        temp += fs[j + 1][k];
                                    }
                                }
                                else break;
                            }
                        }
                    }
                    temp = null;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Ошибка чтения файла конфигурации. " +
                    "Проверьте файл config.cfg. Он должен находиться в папке с программой.");
                Console.WriteLine("Ошибка: " + e.Message);
            }
            return bookmarks;
        }

        private static List<String> ReturnReplacement(string config)
        {
            string temp = null;
            int inter = 0;
            string[] fs = null;
            List<string> replacement = new List<string>();
            try
            {
                fs = File.ReadAllLines(Path.GetFullPath(config));
                //Перебираем конфиг построчно
                for (int i = 0; i < fs.Length; i++)
                {
                    //Встретили закладки
                    if (fs[i] == "[Bookmarks]")
                    {
                        //Начиная с раздела закладки
                        for (int j = i; j < fs.Length; j++)
                        {
                            //Если не в конце файла
                            if ((j + 1) < fs.Length)
                            {
                                //Если не встретили следующий раздел конфига
                                if (fs[j + 1][0] != '/')
                                {
                                    //Перебираем строку
                                    for (int k = 0; k < fs[j + 1].Length; k++)
                                    {
                                        //Если встретили переход от названия закладки к содержимому
                                        if (fs[j + 1][k] == '=')
                                        {
                                            //Запоминаем номер символа перехода
                                            inter = k;
                                        }
                                        //Если не встретили переход, если мы за символом = и если символ не ноль
                                        else if (k > inter && inter != 0)
                                        {
                                            //Посимвольно сохраняем строку ЗАМЕНЫ закладки (после символа =)
                                            temp += fs[j + 1][k];
                                        }
                                    }
                                    replacement.Add((temp).ToString());
                                    temp = null;
                                    inter = 0;
                                }
                                else break;
                            }
                        }
                    }
                    temp = null;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Ошибка чтения файла конфигурации. " +
                    "Проверьте файл config.cfg. Он должен находиться в папке с программой.");
                Console.WriteLine("Ошибка: " + e.Message);
            }
            return replacement;
        }

        private static Word.Range Open(object path)
        {
            Object misObj = Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            app = new Word.Application();
            try
            {
                doc = app.Documents.Add(ref path, ref misObj, ref misObj, ref misObj);
                Word.Range ra = doc.Range();
                Console.WriteLine("Открываем");
                doc.Activate();
                Console.WriteLine("Активируем");
                //Переделать в шаблон
                doc = app.Documents.Open(path);
                doc = app.Documents.Add(ref path, ref misObj, ref misObj, ref misObj);
                return doc.Range();
            }
            catch (Exception exp)
            {
                Close();
                Console.WriteLine("Файл не найден или произошла ошибка чтения файла. Проверьте файл" +
                    "конфигурации.");
                Console.WriteLine("Ошибка: " + exp.Message);
                return null;
            }
        }

        //Ограничение. Замена закладки происходит только в случае оригинальности имени закладки, то есть
        //ИМЕНА закладок не должны повторяться или будет произведена замена по образцу первой закладки
        private static void Booking(Word.Range temppath, List<String> bookMarks, List<String> replaceBooksmarks)
        {
            try
            {
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                int i = 0;
                foreach (Word.Bookmark mark in wBookmarks)
                {
                    Console.WriteLine("Закладка номер " + i.ToString());
                    temppath = mark.Range;
                    if (bookMarks.Contains(mark.Name))
                    {
                        Console.WriteLine("Название закладки: " + mark.Name);
                        temppath.Text = replaceBooksmarks[bookMarks.IndexOf(mark.Name)];
                        Console.WriteLine("Замена закладки: " + temppath.Text);
                    }
                    i++;
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Ошибка чтения закладок: " + exp.Message);
                Close();
                Console.ReadKey();
            }
        }

        //Функция замены слова на другое слово в тексте. Оказалась ненужной
        //private static void Replace(object input, object output)
        //{
        //    try
        //    {
        //        Object missing = Missing.Value;
        //        Console.WriteLine("Input: " + input.ToString());
        //        Console.WriteLine("Output: " + output.ToString());
        //        Word.Range ra = doc.Range();
        //        Console.WriteLine("Содержимое документа: " + ra.Text);
        //        //Word.Range selected = ra;
        //        //selected.Text = "Replacement";
        //        Word.Find fd = app.Selection.Find;
        //        fd.Text = input.ToString();
        //        fd.Replacement.Text = output.ToString();
        //        object replaceAll = Word.WdReplace.wdReplaceAll;
        //        fd.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
        //            ref missing, ref missing, ref missing, ref missing, ref missing,
        //            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        //        Console.WriteLine("Содержимое документа: " + ra.Text);
        //    }
        //    catch (Exception e)
        //    {
        //        Close();
        //        Console.WriteLine("Произошла ошибка замены закладки: " + e.Message);
        //    }
        //}

        private static List<String> onDB()
        {
            //Делаем лист закладок
            List<String> DatabaseBookmarks = new List<string>();
            //Создание соединения с БД
            var dbCon = DBConnection.Instance();
            //Дополнительно переопределяем название БД
            dbCon.DatabaseName = "prestashop";
            //Если соединение произошло
            if (dbCon.IsConnect())
            {
                //ВНИМАНИЕ, ПРОДУМАТЬ НАГРУЖЕННЫЕ СИСТЕМЫ, КОГДА ЗАПИСИ В БД НЕ УСПЕВАЮТ ЗА 
                //ГЕНЕРАЦИЕЙ ТЗ. ВОЗМОЖНО СТОИТ ПЕРЕДАВАТЬ ID КАК ПАРАМЕТР EXEC

                //Запрос к БД
                //string query = "SELECT * FROM phpTest WHERE ID=(SELECT MAX(ID) FROM phpTest)"; // WHERE id=(SELECT MAX(ID) FROM table);";
                string query = "SHOW COLUMNS FROM phpTest;";// WHERE ID=(SELECT MAX(ID) FROM phpTest)";
                //Команда (Послать запрос, к соединению с БД)
                var cmd = new MySqlCommand(query, dbCon.Connection);
                //Ридер получает ответ БД
                var reader = cmd.ExecuteReader();
                
                //Каждая итерация цикла - запись в таблице
                for (int i = 0; i < reader.FieldCount; i++)
                    Console.WriteLine(reader.GetString(i));

                /*while (reader.Read())
                {
                    //Вывести содержимое каждого столбца таблицы
                    for (int i = 1; i < reader.FieldCount; i++)
                        Console.WriteLine(reader.GetString(i));
                            //DatabaseBookmarks.Add((reader.GetString(i)));
                }*/
                //Закрыть соединение с БД
                dbCon.Close();
            }
            return DatabaseBookmarks;
        }

        private static void Close()
        {
            try
            {
                doc.Close();
                app.Quit();
                doc = null;
                app = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Документа не существует, приложение не открылось (возможно уже открыто), или не установлено.");
                Console.WriteLine("Ошибка: " + e.Message);
            }
        }
    }
}
