using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using SheduleSI;
using System.Diagnostics;
using System.Threading;

namespace Rasp
{
    public partial class MainForm : Form
    {
        Repository repository;

        string logs;
        private string facultyLinkBak = "https://portal.esstu.ru/bakalavriat/craspisanEdt.htm";
        private string facultyLinkMag = "https://portal.esstu.ru/spezialitet/craspisanEdt.htm";

        private int threadCount = 6;
        private bool isNotepadRunning;

        private Dictionary<string, SortedDictionary<string, List<List<string>>>> buildingsScheduleMap;
        private List<string> fullClassroomsList;

        Thread thread;

        public MainForm()
        {
            fullClassroomsList = new List<string>();
            isNotepadRunning = false;

            InitializeComponent();

            repository = new Repository();
            this.KeyPreview = true;

            for (int x = 0; x < 13; x++)
            {
                dataGridView1.Rows.Add();
            }
            dataGridView1.Rows[0].Cells[0].Value = "Пнд";
            dataGridView1.Rows[1].Cells[0].Value = "Втр";
            dataGridView1.Rows[2].Cells[0].Value = "Срд";
            dataGridView1.Rows[3].Cells[0].Value = "Чтв";
            dataGridView1.Rows[4].Cells[0].Value = "Птн";
            dataGridView1.Rows[5].Cells[0].Value = "Сбт";

            dataGridView1.Rows[7].Cells[0].Value = "Пнд";
            dataGridView1.Rows[8].Cells[0].Value = "Втр";
            dataGridView1.Rows[9].Cells[0].Value = "Срд";
            dataGridView1.Rows[10].Cells[0].Value = "Чтв";
            dataGridView1.Rows[11].Cells[0].Value = "Птн";
            dataGridView1.Rows[12].Cells[0].Value = "Сбт";
        }

        private async void loadClassroomsSchedule()
        {
            Invoke(new Action(() =>
            {
                allElementsStatus(false);
            }));

            logs = "";
            int progress = 0;
            int errorCount = 0;
            int completedThreads = 0;
            int linksCount = 0;

            buildingsScheduleMap = new Dictionary<string, SortedDictionary<string, List<List<string>>>>
            {
                ["1 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["2 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["3 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["4 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["5 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["6 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["7 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["8 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["9 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["10 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["11 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["12 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["13 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["14 корпус"] = new SortedDictionary<string, List<List<string>>>(),
                ["15 корпус"] = new SortedDictionary<string, List<List<string>>>(),
            };

            async void loadDepartmentPages(List<string> depLinks)
            {
                int localErrorCount = 0;

                ///Загрузка и обработка всех страниц с кафедрами
                foreach (string link in depLinks)
                {
                    try
                    {
                        var splittedDepartmentPage =
                            (await repository.loadDepartmentPage(link))
                                .Replace(" COLOR=\"#0000ff\"", "")
                                .Replace("ff00ff\">", "\a")
                                .Split('\a')
                                .Skip(1);

                        foreach (string teacherSection in splittedDepartmentPage)
                        {
                            string teacherName = "";
                            try
                            {
                                teacherName = Regex.Match(teacherSection, "[а-я]|[А-Я].*</P>").Value.Replace("</P>", "").Trim();
                                //teacherSection.Substring(
                                //teacherSection.IndexOf(Regex(r"[а-я]|[А-Я]")),
                                //teacherSection.indexOf('</P>'));
                            }
                            catch (Exception ex)
                            {
                                logs += ex.Message + new StackTrace(ex, true).GetFrame(0).GetFileLineNumber() + "\n";
                            }

                            var daysOfWeekFromPage =
                                teacherSection.Replace("SIZE=2><P ALIGN=\"CENTER\">", "\a").Split('\a').Skip(1);

                            int j = 0;
                            foreach (string dayOfWeek in daysOfWeekFromPage)
                            {
                                if (j == 12) break;

                                var lessons =
                                    dayOfWeek.Replace("SIZE=1><P ALIGN=\"CENTER\">", "\a").Split('\a').Skip(1);

                                int i = 0;
                                foreach (string lessonSection in lessons)
                                {
                                    if (!lessonSection.Contains("а."))
                                    {
                                        i++;
                                        continue;
                                    }
                                    var fullLesson = lessonSection
                                        .Substring(0, lessonSection.IndexOf("</FONT>"))
                                        .Trim();

                                    var lesson = fullLesson
                                        .Substring(fullLesson.IndexOf("а.") + 2)
                                        .Trim()
                                        .Replace("и/д", "")
                                        .Replace("пр.", "")
                                        .Replace("пр", "")
                                        .Replace("д/кл", "")
                                        .Replace("д/к", "");

                                    var classroom = lesson.Contains(' ')
                                        ? lesson.Substring(0, lesson.IndexOf(' '))
                                        : lesson;

                                    if (!Regex.IsMatch(classroom, "[0-9]"))
                                    //!classroom.Contains(RegExp(r"[0-9]")))
                                    {
                                        i++;
                                        continue;
                                    }

                                    var building = $"{getBuildingByClassroom(classroom)} корпус";
                                    if (!buildingsScheduleMap[building].ContainsKey(classroom))
                                    {
                                        buildingsScheduleMap[building][classroom] = new List<List<string>>(){
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""},
                                            new List<string>(){"", "", "", "", "", ""}
                                        };
                                    }

                                    var finalLesson = $"{teacherName} {fullLesson.Replace(classroom, "")}".Replace('\n', ' ').Replace('\r', ' ');
                                    if (buildingsScheduleMap[building][classroom][j][i].Length <
                                        $"{teacherName} {fullLesson.Replace(classroom, "")}".Length)
                                    {
                                        buildingsScheduleMap[building][classroom][j][i] =
                                            $"{teacherName} {fullLesson.Replace(classroom, "")}";
                                    }

                                    i++;
                                    if (i > 5)
                                    {
                                        break;
                                    }
                                }

                                j++;
                            }
                        }

                        progress++;
                    }
                    catch (Exception ex)
                    {
                        logs += ex.Message + new StackTrace(ex, true).GetFrame(0).GetFileLineNumber() + "\n";

                        localErrorCount++;
                    }

                    if (localErrorCount > 4)
                    {
                        completedThreads++;
                        errorCount += localErrorCount;
                        return;
                    }
                }

                completedThreads++;
            }

            try
            {
                var facultyPages = await repository.loadFacultiesPages(facultyLinkBak, link2: facultyLinkMag);

                //Создания списка ссылок на кафедры
                //
                // Список содержит [threadCount] списков ссылок, которые потом параллельно
                // (ну типо) загружаются и формируют мэп по корпусам
                List<List<string>> departmentLinks = new List<List<string>>();
                for (int iList = 0; iList < threadCount; iList++)
                {
                    departmentLinks.Add(new List<string>());
                }

                List<string> linksList = new List<string>() { "https://portal.esstu.ru/bakalavriat/", "https://portal.esstu.ru/spezialitet/" };
                int i = 0;
                foreach (string facultyPage in facultyPages)
                {
                    List<string> splittedFacultyPage = new List<string>();
                    if (facultyPage.Contains("faculty"))
                    {
                        splittedFacultyPage = Regex.Replace(facultyPage, "<!--.*-->", "")
                            .Replace("href=\"", "\a")
                            .Split('\a')
                            .Skip(1)
                            .ToList();
                    }

                    int j = 0;
                    foreach (string linkSection in splittedFacultyPage)
                    {
                        departmentLinks[j % threadCount].Add(
                          $"{linksList[i]}{linkSection.Substring(0, linkSection.IndexOf('"'))}"
                        );
                        j++;
                    }
                    linksCount += j;
                    i++;
                }

                /// Собственно [threadCount] асинхронных потоков по загрузке страниц. Далее
                /// ождиание окончания их работы с отображением прогресса.
                for (int iList = 0; iList < threadCount; iList++)
                {
                    loadDepartmentPages(departmentLinks[iList]);
                }

                do
                {
                    await Task.Delay(500);

                    Invoke(new Action(() =>
                    {
                        progressBar1.Value = (int)((double)progress / (double)linksCount * 100);
                    }));
                } while (completedThreads < threadCount);

                if (errorCount > 8)
                {
                    Invoke(new Action(() =>
                    {
                        updateButton.Enabled = true;
                    }));
                }
                else
                {
                    var keys = buildingsScheduleMap.Keys.ToList();
                    foreach (string key in keys)
                    {
                        if (buildingsScheduleMap[key].Count < 1)
                        {
                            buildingsScheduleMap.Remove(key);
                        }
                    }

                    fullClassroomsList.Clear();
                    foreach (string building in buildingsScheduleMap.Keys)
                    {
                        fullClassroomsList.AddRange(buildingsScheduleMap[building].Keys);
                    }

                    Invoke(new Action(() =>
                    {
                        buildingComboBox.DataSource = buildingsScheduleMap.Keys.ToList();
                        buildingComboBox.SelectedIndex = 0;

                        allElementsStatus(true);
                    }));
                }
            }
            catch (Exception ex)
            {
                Invoke(new Action(() =>
                {
                    updateButton.Enabled = true;
                }));
                logs += ex.Message + new StackTrace(ex, true).GetFrame(0).GetFileLineNumber() + "\n";
            }

            if (logs.Length > 0)
            {
                using (FileStream fs = new FileStream("./log.txt", FileMode.Create))
                {
                    byte[] buffer = Encoding.Default.GetBytes(logs);
                    await fs.WriteAsync(buffer, 0, buffer.Length);
                    fs.Close();
                    logs = "";
                }

                MessageBox.Show($"Во время загрузки расписания произошли ошибки. Подробности в log.txt", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int getBuildingByClassroom(string classroom)
        {
            if (classroom.StartsWith("0"))
            {
                return 10;
            }
            if (classroom.StartsWith("11"))
            {
                return 11;
            }
            if (classroom.StartsWith("12"))
            {
                return 12;
            }
            if (classroom.StartsWith("13"))
            {
                return 13;
            }
            if (classroom.StartsWith("14"))
            {
                return 14;
            }
            if (classroom.StartsWith("15"))
            {
                return 15;
            }
            if (classroom.StartsWith("2"))
            {
                return 2;
            }
            if (classroom.StartsWith("3"))
            {
                return 3;
            }
            if (classroom.StartsWith("4"))
            {
                return 4;
            }
            if (classroom.StartsWith("5"))
            {
                return 5;
            }
            if (classroom.StartsWith("6"))
            {
                return 6;
            }
            if (classroom.StartsWith("7"))
            {
                return 7;
            }
            if (classroom.StartsWith("8"))
            {
                return 8;
            }
            if (classroom.StartsWith("9"))
            {
                return 9;
            }
            return 1;
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            if (thread.ThreadState != System.Threading.ThreadState.Running) runThread();
        }

        public void showCurrentSchedule()
        {
            //разделитель между неделями в таблице
            int rowDivider = 0;

            for (int x = 0; x < 12; x++)
            {
                if (x > 5) rowDivider = 1;
                for (int y = 0; y < 6; y++)
                {
                    //y + 1 потому что первый столбец - названия дней недели
                    dataGridView1.Rows[x + rowDivider].Cells[y + 1].Value
                        = buildingsScheduleMap[buildingComboBox.SelectedItem.ToString()][classroomComboBox.SelectedItem.ToString()][x][y];
                }
            }
        }

        private async Task saveFiles()
        {
            List<string> classroomsList;
            try
            {
                classroomsList = await repository.loadClassroomsList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка открытия списка аудиторий\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string fileName = Application.StartupPath + "/shabaud.xlsx";
            Excel.Application excelApp = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets;
            Excel.Worksheet sheet;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                //Книга.
                workbooks = excelApp.Workbooks;
            }
            catch (Exception ex)
            {
                if (workbook != null) workbook.Close();
                if (workbooks != null) workbooks.Close();
                if (excelApp != null) excelApp.Quit();

                MessageBox.Show($"Ошибка открытия Excel\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var date = DateTime.Now.ToString("dd.MM.yyyy_HH.mm.ss");
            string outputDir;
            try
            {
                outputDir = Directory.CreateDirectory($"./{date}").FullName;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания папки для сохранения\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool isSuccess = true;
            int progress = 0;

            foreach (string classroom in classroomsList)
            {
                if (!fullClassroomsList.Contains(classroom))
                {
                    progress++;
                    progressBar1.Value = (int)((double)progress / (double)classroomsList.Count * 100);
                    continue;
                }

                try
                {
                    excelApp.Workbooks.Open(fileName);
                    workbook = workbooks[1];
                    //Получаем массив ссылок на листы выбранной книги
                    sheets = workbook.Worksheets;
                    //Выбираем лист 1
                    sheet = (Excel.Worksheet)sheets.get_Item(1);

                    var excelcells = (Excel.Range)sheet.Cells[1, 1];
                    excelcells.Value2 = classroom;

                    int lessonDivider = 0;
                    int dayOfWeekShift = 0;
                    for (int dayOfWeekNum = 2; dayOfWeekNum < 14; dayOfWeekNum++)//столбцы
                    {
                        if (dayOfWeekNum > 7)
                        {
                            lessonDivider = 7;
                            dayOfWeekShift = 6;
                        }
                        else
                        {
                            lessonDivider = 0;
                            dayOfWeekShift = 0;
                        }

                        for (int lessonNum = 3; lessonNum < 9; lessonNum++)
                        {
                            excelcells = (Excel.Range)sheet.Cells[lessonNum + lessonDivider, dayOfWeekNum - dayOfWeekShift];

                            //TODO: изменить десигн шаблона
                            excelcells.Value2 = buildingsScheduleMap[getBuildingByClassroom(classroom) + " корпус"][classroom][dayOfWeekNum - 2][lessonNum - 3];
                            //Aud[selected].getsubject(m - 2 + plus, (n - 3) / 2);
                        }
                    }

                    sheet.SaveAs(outputDir + $"/{classroom.Replace('/', '.')}.xlsx");

                    if (workbook != null) workbook.Close();
                    progress++;
                    progressBar1.Value = (int)((double)progress / (double)classroomsList.Count * 100);
                }
                catch (Exception ex)
                {
                    isSuccess = false;
                    if (workbook != null) workbooks.Close();
                    MessageBox.Show($"Ошибка создания файла расписания\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }
            }

            if (isSuccess)
            {
                MessageBox.Show($"Файлы сохранены в папку {(outputDir.Contains('\\') ? outputDir.Substring(outputDir.LastIndexOf('\\')) : outputDir)}",
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (workbooks != null) workbooks.Close();
            if (excelApp != null) excelApp.Quit();
        }

        private async void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5 && updateButton.Enabled)
            {
                runThread();
            }
            else if (e.KeyCode == Keys.F3)
            {
                if (saveButton.Enabled)
                {
                    allElementsStatus(false);
                    await saveFiles();
                    allElementsStatus(true);
                }
            }
        }

        private void classroomsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            showCurrentSchedule();
        }

        private void allElementsStatus(bool status)
        {
            saveButton.Enabled = status;
            updateButton.Enabled = status;
            buildingComboBox.Enabled = status;
            classroomComboBox.Enabled = status;
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            allElementsStatus(false);
            runThread();
        }

        private void buildingComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            classroomComboBox.DataSource = buildingsScheduleMap[buildingComboBox.SelectedItem.ToString()].Keys.ToList();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.Value != null)
                textBox1.Text = dataGridView1.CurrentCell.Value.ToString();
        }

        private async void saveButton_Click(object sender, EventArgs e)
        {
            allElementsStatus(false);
            await saveFiles();
            allElementsStatus(true);
        }

        private void runThread()
        {
            thread = new Thread(new ThreadStart(loadClassroomsSchedule));
            thread.Start();
        }

        private async void classroomEditButton_Click(object sender, EventArgs e)
        {
            if (isNotepadRunning) return;
            isNotepadRunning = true;

            await Task.Run(new Action(() =>
            {
                try
                {
                    using (Process pProcess = new Process())
                    {
                        pProcess.StartInfo.FileName = @"notepad";
                        pProcess.StartInfo.Arguments = Application.StartupPath + "/classrooms.txt";
                        pProcess.Start();
                        pProcess.WaitForExit();

                        isNotepadRunning = false;
                    }
                }
                catch(Exception ex)
                {
                    isNotepadRunning = false;
                    MessageBox.Show($"Ошибка открытия блокнота\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }));
        }
    }
}
