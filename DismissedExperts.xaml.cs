using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using static System.Net.Mime.MediaTypeNames;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для DismissedExperts.xaml
    /// </summary>
    public partial class DismissedExperts : Page
    {
        public Experts _expert;
        public string logi;
        public string rol;
        public DismissedExperts(Experts expert)
        {
            InitializeComponent();
            var currentExperts = ExpertiseEntities.GetContext().Experts.Where(p => p.Status == "Уволен").ToList();
            ExpertsListView.ItemsSource = currentExperts;
            ima.Text = expert.Surname + " " + expert.Name + " " + expert.Patronymic;
            FiltrExpert.SelectedIndex = 0;
            role.Text = "Администратор";
            Update();
        }




        private void Update()
        {
            var currentExperts = ExpertiseEntities.GetContext().Experts.Where(p => p.Status == "Уволен").ToList();

            if (FiltrExpert.SelectedIndex == 1)
            {
                currentExperts = currentExperts.Where(p => p.Expertise.Count >= 0 && p.Expertise.Count <= 5).ToList();
            }
            if (FiltrExpert.SelectedIndex == 2)
            {
                currentExperts = currentExperts.Where(p => p.Expertise.Count > 5 && p.Expertise.Count <= 10).ToList();
            }
            if (FiltrExpert.SelectedIndex == 3)
            {
                currentExperts = currentExperts.Where(p => p.Expertise.Count > 10).ToList();
            }


            currentExperts = currentExperts.Where(p => p.FIO.ToLower().Contains(PoiskPoFIO.Text.ToLower())).ToList();
            ExpertsListView.ItemsSource = currentExperts.ToList();

            // Сортировка
            if (New.IsChecked == true)
            {
                currentExperts = currentExperts.OrderBy(p => p.Surname).ToList();
            }
            else if (Old.IsChecked == true)
            {
                currentExperts = currentExperts.OrderByDescending(p => p.Surname).ToList();

            }
            else
            {
                // Сортировка по ID (как при запуске)
                currentExperts = currentExperts.OrderBy(p => p.ID_Expert).ToList();
            }

            ExpertsListView.ItemsSource = currentExperts;


            int totalRecords = ExpertiseEntities.GetContext().Experts.Where(p => p.Status == "Уволен").Count();
            int filteredRecords = currentExperts.Where(p => p.Status == "Уволен").Count();
            Count.Text = "количество " + filteredRecords + " из " + totalRecords;
        }

        // Добавьте этот метод в класс ExpertPage
        private void ExportReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Запрашиваем год для отчета
                var yearDialog = new YearSelectionDialog();
                yearDialog.Owner = Window.GetWindow(this);

                if (yearDialog.ShowDialog() != true)
                {
                    return; // Пользователь отменил выбор
                }

                int selectedYear = yearDialog.SelectedYear;

                // Получаем уволенных сотрудников, кроме администраторов
                var context = ExpertiseEntities.GetContext();
                var employees = context.Experts
                    .Where(p => p.ID_Post != 4 && p.Status == "Уволен")
                    .ToList();

                if (employees.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта", "Внимание",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Формируем имя файла
                string fileName = $"Отчет_по_уволенным_сотрудникам_{selectedYear}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                // Создаем папку Export в корне проекта
                string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                string exportFolder = System.IO.Path.Combine(projectRoot, "Export");

                if (!Directory.Exists(exportFolder))
                {
                    Directory.CreateDirectory(exportFolder);
                }

                string filePath = System.IO.Path.Combine(exportFolder, fileName);

                // Проверяем существование файла
                int fileNumber = 1;
                string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(fileName);
                string extension = System.IO.Path.GetExtension(fileName);

                while (File.Exists(filePath))
                {
                    string newFileName = $"{fileNameWithoutExtension}_{fileNumber}{extension}";
                    filePath = System.IO.Path.Combine(exportFolder, newFileName);
                    fileNumber++;
                }

                // Экспортируем данные
                ExportEmployeeReportToExcel(employees, filePath, selectedYear);

                MessageBox.Show($"Отчет за {selectedYear} год успешно экспортирован!\n\nФайл: {filePath}",
                    "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте отчета: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportEmployeeReportToExcel(List<Experts> employees, string filePath, int year)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = $"Отчет_уволенные_{year}";

                // Определяем текущую дату
                DateTime currentDate = DateTime.Now;
                int currentMonth = currentDate.Month;
                int currentYear = currentDate.Year;

                // Определяем максимальное количество месяцев для вывода
                int maxMonth = (year == currentYear) ? currentMonth : 12;

                // Массив названий месяцев
                string[] monthNames = { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };

                // Заголовки столбцов
                string[] baseHeaders = { "№ п/п", "ФИО сотрудника", "Должность", "Оклад (руб./мес)", "Дата увольнения" };
                string[] monthHeaders = { "Кол-во экспертиз", "Премия (10%)", "Итого ЗП за месяц" };
                string[] totalHeaders = { "Всего экспертиз", "Общая премия", "Итоговая ЗП" };

                int currentCol = 1;

                // Базовые заголовки
                foreach (var header in baseHeaders)
                {
                    worksheet.Cells[1, currentCol] = header;
                    currentCol++;
                }

                // Заголовки по месяцам (только для нужных месяцев)
                int monthStartCol = currentCol;
                for (int month = 0; month < maxMonth; month++)
                {
                    worksheet.Cells[1, currentCol] = monthNames[month];
                    worksheet.Cells[1, currentCol + 1] = "";
                    worksheet.Cells[1, currentCol + 2] = "";

                    Excel.Range mergeRange = worksheet.Range[worksheet.Cells[1, currentCol], worksheet.Cells[1, currentCol + 2]];
                    mergeRange.Merge();
                    mergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    mergeRange.Font.Bold = true;

                    currentCol += 3;
                }

                // Подзаголовки для месяцев
                currentCol = monthStartCol;
                for (int month = 0; month < maxMonth; month++)
                {
                    for (int i = 0; i < monthHeaders.Length; i++)
                    {
                        worksheet.Cells[2, currentCol + i] = monthHeaders[i];
                    }
                    currentCol += 3;
                }

                // Итоговые заголовки
                int totalStartCol = currentCol;
                worksheet.Cells[1, totalStartCol] = "ИТОГО";
                worksheet.Cells[1, totalStartCol + 1] = "";
                worksheet.Cells[1, totalStartCol + 2] = "";

                Excel.Range totalMergeRange = worksheet.Range[worksheet.Cells[1, totalStartCol], worksheet.Cells[1, totalStartCol + 2]];
                totalMergeRange.Merge();
                totalMergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                totalMergeRange.Font.Bold = true;

                for (int i = 0; i < totalHeaders.Length; i++)
                {
                    worksheet.Cells[2, totalStartCol + i] = totalHeaders[i];
                }

                // Форматирование заголовков
                Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, totalStartCol + 2]];
                headerRange.Font.Bold = true;
                headerRange.Font.Size = 11;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                headerRange.WrapText = true;

                // Заполняем данные по сотрудникам
                int row = 3;
                int serialNumber = 1;

                // Для итогов по всем сотрудникам
                decimal[] totalExpertiseAll = new decimal[maxMonth];
                decimal[] totalBonusAll = new decimal[maxMonth];
                decimal[] totalSalaryAll = new decimal[maxMonth];

                decimal grandTotalExpertise = 0;
                decimal grandTotalBonus = 0;
                decimal grandTotalSalary = 0;

                foreach (var employee in employees)
                {
                    // Определяем месяц увольнения (для уволенных сотрудников)
                    int dismissalMonth = 13;
                    if (employee.Date_Dismissed.HasValue)
                    {
                        if (employee.Date_Dismissed.Value.Year < year)
                        {
                            dismissalMonth = 0; // Уволен до начала отчетного года
                        }
                        else if (employee.Date_Dismissed.Value.Year == year)
                        {
                            dismissalMonth = employee.Date_Dismissed.Value.Month;
                        }
                    }

                    // Если сотрудник уволен до начала отчетного года - пропускаем
                    if (dismissalMonth == 0) continue;

                    // Получаем экспертизы сотрудника за указанный год (только до месяца увольнения)
                    var expertises = employee.Expertise
                        .Where(e => e.Date_Start.HasValue && e.Date_Start.Value.Year == year)
                        .ToList();

                    decimal monthlySalary = employee.Post?.Salary ?? 0;

                    // Базовые данные
                    worksheet.Cells[row, 1] = serialNumber++;
                    worksheet.Cells[row, 2] = employee.FIO;
                    worksheet.Cells[row, 3] = employee.Post?.Post_Name ?? "Не указана";
                    worksheet.Cells[row, 4] = monthlySalary;
                    worksheet.Cells[row, 5] = employee.Date_Dismissed?.ToString("dd.MM.yyyy") ?? "Дата не указана";

                    int currentDataCol = monthStartCol;
                    decimal employeeTotalExpertise = 0;
                    decimal employeeTotalBonus = 0;
                    decimal employeeTotalSalary = 0;

                    // Данные по месяцам (только до месяца увольнения)
                    int maxMonthForEmployee = Math.Min(maxMonth, dismissalMonth == 13 ? maxMonth : dismissalMonth);

                    for (int month = 1; month <= maxMonthForEmployee; month++)
                    {
                        var monthExpertises = expertises
                            .Where(e => e.Date_Start.Value.Month == month)
                            .ToList();

                        int monthCount = monthExpertises.Count;
                        decimal monthBonus = monthExpertises.Sum(e => e.Cost) * 0.10m;
                        decimal monthSalary = monthlySalary + monthBonus;

                        worksheet.Cells[row, currentDataCol] = monthCount;
                        worksheet.Cells[row, currentDataCol + 1] = monthBonus;
                        worksheet.Cells[row, currentDataCol + 2] = monthSalary;

                        employeeTotalExpertise += monthCount;
                        employeeTotalBonus += monthBonus;
                        employeeTotalSalary += monthSalary;

                        totalExpertiseAll[month - 1] += monthCount;
                        totalBonusAll[month - 1] += monthBonus;
                        totalSalaryAll[month - 1] += monthSalary;

                        currentDataCol += 3;
                    }

                    // Пропускаем оставшиеся ячейки, если сотрудник уволен
                    if (dismissalMonth <= maxMonth)
                    {
                        int skippedMonths = maxMonth - dismissalMonth;
                        currentDataCol += skippedMonths * 3;
                    }

                    // Итоги за период
                    worksheet.Cells[row, totalStartCol] = employeeTotalExpertise;
                    worksheet.Cells[row, totalStartCol + 1] = employeeTotalBonus;
                    worksheet.Cells[row, totalStartCol + 2] = employeeTotalSalary;

                    grandTotalExpertise += employeeTotalExpertise;
                    grandTotalBonus += employeeTotalBonus;
                    grandTotalSalary += employeeTotalSalary;

                    row++;
                }

                // Строка с итогами по месяцам
                if (row > 3)
                {
                    row++;
                    worksheet.Cells[row, 1] = "";
                    worksheet.Cells[row, 2] = "ИТОГО ЗА МЕСЯЦ:";
                    worksheet.Cells[row, 3] = "";
                    worksheet.Cells[row, 4] = "";
                    worksheet.Cells[row, 5] = "";

                    int currentDataColTotal = monthStartCol;
                    for (int month = 0; month < maxMonth; month++)
                    {
                        worksheet.Cells[row, currentDataColTotal] = totalExpertiseAll[month];
                        worksheet.Cells[row, currentDataColTotal + 1] = totalBonusAll[month];
                        worksheet.Cells[row, currentDataColTotal + 2] = totalSalaryAll[month];
                        currentDataColTotal += 3;
                    }

                    worksheet.Cells[row, totalStartCol] = grandTotalExpertise;
                    worksheet.Cells[row, totalStartCol + 1] = grandTotalBonus;
                    worksheet.Cells[row, totalStartCol + 2] = grandTotalSalary;

                    Excel.Range totalRow = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, totalStartCol + 2]];
                    totalRow.Font.Bold = true;
                    totalRow.Interior.Color = Excel.XlRgbColor.rgbLightYellow;
                }

                // Форматирование числовых ячеек
                if (row > 3)
                {
                    Excel.Range countRange = worksheet.Range[
                        worksheet.Cells[3, monthStartCol],
                        worksheet.Cells[row, totalStartCol + 2]];
                    countRange.NumberFormat = "#,##0";

                    Excel.Range moneyRange = worksheet.Range[
                        worksheet.Cells[3, monthStartCol + 1],
                        worksheet.Cells[row, totalStartCol + 2]];
                    moneyRange.NumberFormat = "#,##0.00";
                }

                // Границы
                Excel.Range dataRange = worksheet.Range[
                    worksheet.Cells[1, 1],
                    worksheet.Cells[row, totalStartCol + 2]];
                dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                worksheet.Columns.AutoFit();
                worksheet.Columns[2].ColumnWidth = 25;
                worksheet.Columns[4].ColumnWidth = 15;
                worksheet.Columns[5].ColumnWidth = 15;

                // Закрепление строк
                Excel.Range freezeRange = worksheet.Range["A3"];
                freezeRange.Select();
                excelApp.ActiveWindow.FreezePanes = true;

                workbook.SaveAs(filePath);
                workbook.Close();
                excelApp.Quit();
            }
            finally
            {
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }        /// <summary>
                 /// Экспорт отчетности по сотрудникам в Excel
                 /// </summary>

                 /// <summary>
                 /// Экспорт отчета по сотрудникам в Excel
                 /// </summary>
        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            Manager.MainFrame.Navigate(new AddEditExperts((sender as Button).DataContext as Experts));
            Update();
        }

        private void FiltrExpert_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Update();
        }


        private void PoiskPoFIO_TextChanged(object sender, TextChangedEventArgs e)
        {
            Update();
        }

        private void New_Checked(object sender, RoutedEventArgs e)
        {
            // При выборе "Новые" снимаем выделение со "Старые"
            Old.IsChecked = false;
            Update();
        }

        private void New_Unchecked(object sender, RoutedEventArgs e)
        {
            // При снятии выделения с "Новые" показываем по ID
            Update();
        }

        private void Old_Checked(object sender, RoutedEventArgs e)
        {
            // При выборе "Старые" снимаем выделение с "Новые"
            New.IsChecked = false;
            Update();
        }

        private void Old_Unchecked(object sender, RoutedEventArgs e)
        {
            // При снятии выделения с "Старые" показываем по ID
            Update();
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Manager.MainFrame.Navigate(new AddEditExperts(null));
            Update();
        }

        private void ExpertsListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Update();
        }

        private void ExpertsListView_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            Update();
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            Update();
        }
    }
}
