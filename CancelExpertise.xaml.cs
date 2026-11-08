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

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для CancelExpertise.xaml
    /// </summary>
    public partial class CancelExpertise : Page
    {
        public Experts _expert;
        public string logi;
        public string rol;
        public CancelExpertise(Experts expert)
        {
            InitializeComponent();
            _expert = expert;
            var currentExpertise = ExpertiseEntities.GetContext().Expertise.ToList();
            ExpertiseListView.ItemsSource = currentExpertise;
            if (expert != null)
            {
                logi = expert.Surname + " " + expert.Name + " " + expert.Patronymic;
                ima.Text = logi;
                switch (expert.ID_Post)
                {
                    case 1:
                        rol = "Главный государственный судебный эксперт";
                        break;
                    case 2:
                        rol = "Старший государственнй суд эксперт";
                        break;
                    case 3:
                        rol = "Государственный судебный эксперт";
                        break;
                    case 4:
                        rol = "Администратор";
                        break;
                    default:
                        rol = "Гость";
                        break;

                }
            }
            else
            {
                rol = "Гость";

            }
            FiltrKomnat.SelectedIndex = 0;
            role.Text = rol;
            Update();
        }

        private void Update()
        {
            var currentExpertise = ExpertiseEntities.GetContext().Expertise.Where(p => p.Status == "Отменена").ToList();

            if (_expert != null) // ID_Post = 4 это администратор
            {
                currentExpertise = currentExpertise.Where(p => p.ID_Expert == _expert.ID_Expert).ToList();
            }
            int totalRecords = currentExpertise.Count;
            if (FiltrKomnat.SelectedIndex == 1)
            {
                currentExpertise = currentExpertise.Where(p => p.Purpose1 == 1).ToList();
            }
            if (FiltrKomnat.SelectedIndex == 2)
            {
                currentExpertise = currentExpertise.Where(p => p.Purpose1 == 2).ToList();
            }
            if (FiltrKomnat.SelectedIndex == 3)
            {
                currentExpertise = currentExpertise.Where(p => p.Purpose1 >= 3).ToList();
            }


            currentExpertise = currentExpertise.Where(p => p.Address1.ToLower().Contains(PoiskPoAdress.Text.ToLower())).ToList();
            ExpertiseListView.ItemsSource = currentExpertise.ToList();
            if (New.IsChecked.Value)
            {
                currentExpertise = currentExpertise.OrderByDescending(p => p.Date_Start.HasValue ? p.Date_Start.Value.Date : DateTime.MinValue)
                                          .ThenByDescending(p => p.Date_Start) // для сортировки по времени, если даты одинаковые
                                          .ToList();
            }
            if (Old.IsChecked.Value)
            {
                currentExpertise = currentExpertise.OrderBy(p => p.Date_Start.HasValue ? p.Date_Start.Value.Date : DateTime.MaxValue)
                                          .ThenBy(p => p.Date_Start) // для сортировки по времени, если даты одинаковые
                                          .ToList();
            }

            ExpertiseListView.ItemsSource = currentExpertise;


            int filteredRecords = currentExpertise.Count;
            Count.Text = "количество " + filteredRecords + " из " + totalRecords;
        }


        private void PoiskPoAdress_TextChanged(object sender, TextChangedEventArgs e)
        {
            Update();
        }

        private void New_Checked(object sender, RoutedEventArgs e)
        {
            Update();
        }

        private void Old_Checked(object sender, RoutedEventArgs e)
        {
            Update();
        }

        private void FiltrKomnat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Update();
        }

        private void ExpertiseListView_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            Update();
        }



        // Измененный метод экспорта - вызывает оба экспорта
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем кнопку, по которой кликнули
                Button button = sender as Button;
                if (button == null)
                {
                    MessageBox.Show("Ошибка получения данных", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Получаем экспертизу из DataContext кнопки
                Expertise selectedExpertise = button.DataContext as Expertise;

                if (selectedExpertise == null)
                {
                    MessageBox.Show("Выберите экспертизу для экспорта", "Внимание",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }


                // Экспорт в Word
                ExportToWord(selectedExpertise);

                MessageBox.Show("Данные успешно экспортированы в Word!",
                    "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Метод для экспорта в Word
        private void ExportToWord(Expertise selectedExpertise)
        {
            // Путь к шаблону
            string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string templatePath = System.IO.Path.Combine(projectRoot, "Export", "Шаблон отмена.docx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Шаблон не найден по пути: {templatePath}");
            }

            // Формируем имя выходного файла
            string fileName = GenerateFileName(selectedExpertise.Realties) + "_Отмена.docx";
            string exportFolder = System.IO.Path.Combine(projectRoot, "Export");

            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
            }

            string outputPath = System.IO.Path.Combine(exportFolder, fileName);

            // Проверяем существование файла и генерируем уникальное имя
            int fileNumber = 1;
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(fileName);
            string extension = System.IO.Path.GetExtension(fileName);

            while (File.Exists(outputPath))
            {
                string newFileName = $"{fileNameWithoutExtension}_{fileNumber}{extension}";
                outputPath = System.IO.Path.Combine(exportFolder, newFileName);
                fileNumber++;
            }

            // Копируем шаблон в новый файл
            File.Copy(templatePath, outputPath, true);

            // Заполняем Word документ
            FillWordDocument(outputPath, selectedExpertise);
        }

        // Метод для заполнения Word документа
        private void FillWordDocument(string filePath, Expertise expertise)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                doc = wordApp.Documents.Open(filePath);

                // Получаем данные для замены
                var replacements = GetWordReplacements(expertise);

                // Выполняем замену
                foreach (var replacement in replacements)
                {
                    ReplaceTextInWordDocument(doc, replacement.Key, replacement.Value);
                }

                // Сохраняем и закрываем
                doc.Save();
                doc.Close();
                wordApp.Quit();
            }
            finally
            {
                if (doc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                if (wordApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Метод для получения словаря замен для Word
        private Dictionary<string, string> GetWordReplacements(Expertise expertise)
        {
            var realty = expertise.Realties;
            var caseInfo = expertise.Case;
            var expert = expertise.Experts;

            // Форматируем даты
            string dateStart = expertise.Date_Start?.ToString("dd.MM.yyyy") ?? "_____";
            string dateEnd = expertise.Date_End?.ToString("dd.MM.yyyy") ?? "_____";
            string dateCase = caseInfo?.Date_Case.ToString("dd.MM.yyyy") ?? "_____";

            // Формируем полный адрес
            string fullAddress = realty?.FullAddress ?? "Адрес не указан";

            // Формируем информацию о комнатах
            string purposeText = "";
            if (realty?.Purpose != null)
            {
                int rooms = realty.Purpose;
                purposeText = rooms == 1 ? "1-комнатная" :
                              rooms == 2 ? "2-комнатная" :
                              rooms == 3 ? "3-комнатная" : $"{rooms}-комнатная";
            }

            // Формируем материал стен
            string wallMaterial = GetWallMaterialForWord(realty?.Wall_Material);

            // Формируем отделку
            string finishing = GetFinishingForWord(realty?.Finishing);

            // Формируем ФИО эксперта
            string expertFIO = expert?.FIO ?? "_________________________";

            // Формируем номер уголовного дела
            string caseNumber = caseInfo?.Number_Case ?? "_____";

            // Формируем суд
            string court = caseInfo?.court ?? "______________";

            // Формируем цену с округлением
            string price = realty?.Price.HasValue == true
                ? $"{Math.Round(realty.Price.Value, 0):N0}" : "_________";

            return new Dictionary<string, string>
            {
                { "[ID_Case]", caseNumber },
                { "[Date_Case]", dateCase },
                { "[Date_Start]", dateStart },
                { "[Date_End]", dateEnd },
                { "[court]", court },
                { "[Expert]", expertFIO },
                { "[Address1]", fullAddress },
                { "[Floor]", realty?.Floor.ToString() ?? "___" },
                { "[Area]", realty?.Area.ToString("N1") ?? "___" },
                { "[Purpose]", purposeText },
                { "[Finishing]", finishing },
                { "[Wall_Material]", wallMaterial },
                { "[Price]", price }
            };
        }

        // Метод для замены текста в Word документе
        private void ReplaceTextInWordDocument(Word.Document doc, string findText, string replaceText)
        {
            if (string.IsNullOrEmpty(findText)) return;

            Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Replacement.ClearFormatting();

            object findObj = findText;
            object replaceObj = replaceText;
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildcards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object wrap = Word.WdFindWrap.wdFindContinue;
            object format = false;
            object replace = Word.WdReplace.wdReplaceAll;

            range.Find.Execute(ref findObj, ref matchCase, ref matchWholeWord, ref matchWildcards,
                ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format,
                ref replaceObj, ref replace, ref format, ref format, ref format);
        }

        // Метод для генерации имени файла на основе адреса
        private string GenerateFileName(Realties realty)
        {
            
            string address = "";

            // Используем NameStreet (вычисляемое свойство, которое возвращает название улицы)
            if (!string.IsNullOrWhiteSpace(realty.NameStreet))
                address += realty.NameStreet;

            if (!string.IsNullOrWhiteSpace(realty.House))
                address += "_" + realty.House;

            if (!string.IsNullOrWhiteSpace(realty.Apartment_number))
                address += "_кв" + realty.Apartment_number;

            // Очищаем от недопустимых символов
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                address = address.Replace(c, '_');
            }

            address = address.Replace(" ", "_");

            return string.IsNullOrWhiteSpace(address)
                ? $"Экспертиза_{DateTime.Now:yyyyMMdd_HHmmss}"
                : $"Экспертиза_{address}";
        }

        
        // Вспомогательные методы для Word
        private string GetWallMaterialForWord(string material)
        {
            if (string.IsNullOrEmpty(material)) return "Данные отсутствуют";

            switch (material.ToLower())
            {
                case "кирпич":
                    return "кирпичные";
                case "панель":
                    return "панельные";
                case "монолит":
                    return "монолитные";
                case "дерево":
                    return "деревянные";
                default:
                    return material;
            }
        }

        private string GetFinishingForWord(string finishing)
        {
            if (string.IsNullOrEmpty(finishing)) return "Данные отсутствуют";

            switch (finishing.ToLower())
            {
                case "черновая":
                    return "черновая отделка";
                case "чистовая":
                    return "чистовая отделка";
                case "евро":
                    return "евроремонт";
                case "косметический":
                    return "косметический ремонт";
                default:
                    return finishing;
            }
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            Update();
        }
    }
}
