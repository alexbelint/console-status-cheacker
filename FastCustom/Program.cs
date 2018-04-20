using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace FastCustom
{
    class Program
    {
        static void Main(string[] args)
        {
            {
                Run();
            }
        }
        public static string inputFileName = string.Empty;
        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]

        public static void Run()
        {
            //Console.WriteLine("Console is working now...");
            // Create a new FileSystemWatcher and set its properties.
            FileSystemWatcher watcher = new FileSystemWatcher();
            //watcher.Path = args[1];
            watcher.Path = ConfigurationSettings.AppSettings["RootFolder"];
            /* Watch for changes in LastAccess and LastWrite times, and
               the renaming of files or directories. */
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
               | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            // Only watch text files.
            watcher.Filter = "*.xls*";

            // Add event handlers.

            // Begin watching.
            watcher.EnableRaisingEvents = true;
            foreach (string file in Directory.EnumerateFiles(ConfigurationSettings.AppSettings["RootFolder"], "*.xls"))
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                inputFileName = Path.GetFileNameWithoutExtension(file);
                ArrayList vagonsList = new ArrayList();

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application(); //создаем COM-объект Excel
                excel.Visible = false; //делаем объект невидимым
                excel.Workbooks.Add(Type.Missing); //добавляем книгу
                excel.SheetsInNewWorkbook = 1;//количество листов в книге
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks[1]; //получам ссылку на первую открытую книгу
                Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Worksheets.get_Item(1);//получаем ссылку на первый лист
                sheet.Name = "Report" + " " + DateTime.Now.ToString("dd.MM.yy");
                sheet.Cells[1, 1] = "№ КОНТЕЙНЕРА";
                sheet.Cells[1, 2] = "СТАНЦ";
                sheet.Cells[1, 3] = "ОПЕР";
                sheet.Cells[1, 4] = "ДАТА";
                sheet.Cells[1, 5] = "ВРЕМЯ";
                sheet.Cells[1, 6] = "СОСТ";
                sheet.Cells[1, 7] = "N ОТПР";
                sheet.Cells[1, 8] = "N ВАГОНА";

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (j == 1)
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        //add useful things here!  
                        //    string.Format("(:217 0:1680 {0}:)", xlRange.Cells[i, j].Value2.ToString()));
                        vagonsList.Add(xlRange.Cells[i, j].Value2.ToString());
                    }
                }
                for (int i = 0; i < vagonsList.Count; i++)
                {
                    int list = vagonsList.Count;
                    var container = vagonsList[i];
                    if (vagonsList != null)
                    {
                        File.WriteAllText(ConfigurationManager.AppSettings["querryFolder"] + "01" + 11 + "000" + ".000", string.Format("(:217 0:1680 {0}:)", container));
                        Console.WriteLine("Now processing: {0}", container);
                        System.Threading.Thread.Sleep(9500);
                    }
                    else
                    {
                        break;
                    }
                    string answerFileName = "01" + 11 + "2400";
                    //string answerFileName = "01" + 15 + "2400"; //15й порт Левитан, 11 мой
                    DirectoryInfo dirConcentratorPath = new DirectoryInfo(ConfigurationManager.AppSettings["answerFolder"]);
                    FileInfo[] fileInDir = dirConcentratorPath.GetFiles(answerFileName + "*.*");
                    foreach (FileInfo foundFile in fileInDir)
                    {
                        string fullName = foundFile.FullName;
                        var lines = File.ReadAllLines(foundFile.FullName, Encoding.GetEncoding(866));
                        sheet.get_Range("E2", string.Format("E{0}", vagonsList.Count)).NumberFormat = "@";
                        string text = ""; // переменная для поиска ключа в файле
                        using (StreamReader sr = new StreamReader(foundFile.FullName, Encoding.GetEncoding(866)))
                        {
                            text = sr.ReadToEnd();
                            Regex regexNoInfo = new Regex("[Н][Е][Т]\\s[А-Я]{10}");
                            foreach (var line in lines)
                            {
                                var NoInfoMatches = regexNoInfo.Matches(line);
                                if (NoInfoMatches.Count > 0)
                                {
                                    var res = NoInfoMatches[0].Value;
                                    string noinfo = "-";
                                    sheet.Cells[i + 2, 1].Value = container;
                                    sheet.Cells[i + 2, 2].Value = noinfo;
                                    sheet.Cells[i + 2, 3].Value = noinfo;
                                    sheet.Cells[i + 2, 4].Value = noinfo;
                                    sheet.Cells[i + 2, 5].Value = noinfo;
                                    sheet.Cells[i + 2, 6].Value = noinfo;
                                    sheet.Cells[i + 2, 7].Value = noinfo;
                                    sheet.Cells[i + 2, 8].Value = noinfo;
                                    //sheet.Cells[i + 1, 9].Value = noinfo;
                                    break;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            Regex regexOkInfo = new Regex(@"([О][П][Е][Р][А][Ц][И][И]\s[С]\s[К])");
                            foreach (var line in lines)
                            {
                                var okInfoMatches = regexOkInfo.Matches(line);
                                if (okInfoMatches.Count > 0)
                                {
                                    string tempAnswer = lines[lines.Length - 4];
                                    #region using regular expression
                                    string pattern = @"([A-Я]{4})|([А-Я]{3}.)|(\d\d[.]\d\d[.]\d\d)|([А-Я]{3}.)|(\d{2}[-]\d{2})|(\d{8})|(\d{6}?)";
                                    Regex rgx = new Regex(pattern);
                                    MatchCollection matchList = Regex.Matches(tempAnswer, pattern);
                                    var results = matchList.Cast<Match>().Select(match => match.Value).ToList();
                                    var station = matchList[0].Value;
                                    var operation = matchList[1].Value;
                                    var dt = matchList[2].Value;
                                    DateTime dateOfOperation = Convert.ToDateTime(dt);
                                    string timeOfOperation = matchList[3].Value;
                                    var state = matchList[4].Value;
                                    var otpravka = matchList[5].Value;
                                    if (results.Count == 6)
                                    {
                                        sheet.Cells[i + 2, 1].Value = container; //выводим в столбик название контейнеров
                                        sheet.Cells[i + 2, 2].Value = station;
                                        sheet.Cells[i + 2, 3].Value = operation;
                                        sheet.Cells[i + 2, 4].Value = dateOfOperation;
                                        sheet.Cells[i + 2, 5].Value = timeOfOperation;
                                        sheet.Cells[i + 2, 6].Value = state;
                                        sheet.Cells[i + 2, 7].Value = otpravka;
                                        string noinfo = "-";
                                        sheet.Cells[i + 1, 8].Value = noinfo;
                                        //sheet.Cells[i + 1, 9].Value = noinfo;
                                    }
                                    else
                                    {
                                        var vagon = matchList[6].Value;
                                        //var index = matchList[7].Value;
                                        //sheet.get_Range("E2", string.Format("E{0}", vagonsList.Count)).NumberFormat = "@";
                                        #endregion
                                        sheet.Cells[i + 2, 1].Value = container; //выводим в столбик название контейнеров
                                        sheet.Cells[i + 2, 2].Value = station;
                                        sheet.Cells[i + 2, 3].Value = operation;
                                        sheet.Cells[i + 2, 4].Value = dateOfOperation;
                                        sheet.Cells[i + 2, 5].Value = timeOfOperation;
                                        sheet.Cells[i + 2, 6].Value = state;
                                        sheet.Cells[i + 2, 7].Value = otpravka;
                                        sheet.Cells[i + 2, 8].Value = vagon;

                                    }
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        File.Delete(foundFile.FullName);
                    }
                }
                #region Formatting excel
                var formatTable = vagonsList.Count + 1; // костыль, чтобы последняя строка тоже получила форматирование
                sheet.get_Range("B1", string.Format("H{0}", formatTable)).Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; //выравнивание по центру
                                                                                                                                            //sheet.get_Range("E2", string.Format("E{0}", col2Items.Count)).NumberFormat = "hh";
    

                Microsoft.Office.Interop.Excel.Range chartRange;
                chartRange = sheet.get_Range("a1", "h1");
                foreach (Microsoft.Office.Interop.Excel.Range cells in chartRange.Cells)
                {
                    cells.BorderAround2();
                }
                #endregion
                try
                {
                    sheet.Columns.AutoFit(); // autofit
                    string reportName = inputFileName + " от " + DateTime.Now.ToString("dd.MM.yyyy, HH-mm");
                    workbook.SaveAs(ConfigurationManager.AppSettings["reportFolder"] + reportName + ".xlsx");
                    //workbook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + reportName + ".xlsx");
                    excel.Workbooks.Close();
                    excel.Quit();
                    Console.WriteLine("Processing file {0} is completed.", inputFileName);

                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    //errors with saving
                }
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                //Console.WriteLine("Press \'q\' to quit the console.");
                //while (Console.Read() != 'q') ;
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}