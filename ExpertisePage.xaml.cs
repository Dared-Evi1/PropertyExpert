// Добавьте в начало файла using:
using Microsoft.Win32;
using System;
using System.Collections;
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
    /// Логика взаимодействия для ExpertisePage.xaml
    /// </summary>
    public partial class ExpertisePage : Page
    {
        public Experts _expert;
        public string logi;
        public string rol;

        public ExpertisePage(Experts expert)
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
                }
            }
            FiltrKomnat.SelectedIndex = 0;
            role.Text = rol;
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

        private void Update()
        {

            var currentExpertise = ExpertiseEntities.GetContext().Expertise
                .Where(p => p.Status == "В процессе" || p.Status == "Завершена")
                .ToList();

            if (_expert != null)
            {
                currentExpertise = currentExpertise.Where(p => p.ID_Expert == _expert.ID_Expert).ToList();
            }

            int totalRecords = currentExpertise.Count;

            // Фильтр по комнатам
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

            // Поиск по адресу
            currentExpertise = currentExpertise
                .Where(p => p.Address1.ToLower().Contains(PoiskPoAdress.Text.ToLower()))
                .ToList();

            // Сортировка
            if (New.IsChecked == true)
            {
                // По убыванию даты (сначала новые)
                currentExpertise = currentExpertise
                    .OrderByDescending(p => p.Date_Start.HasValue ? p.Date_Start.Value.Date : DateTime.MinValue)
                    .ThenByDescending(p => p.Date_Start)
                    .ToList();
            }
            else if (Old.IsChecked == true)
            {
                // По возрастанию даты (сначала старые)
                currentExpertise = currentExpertise
                    .OrderBy(p => p.Date_Start.HasValue ? p.Date_Start.Value.Date : DateTime.MaxValue)
                    .ThenBy(p => p.Date_Start)
                    .ToList();
            }
            else
            {
                // Сортировка по ID (как при запуске)
                currentExpertise = currentExpertise.OrderBy(p => p.ID_Expertise).ToList();
            }

            ExpertiseListView.ItemsSource = currentExpertise;

            int filteredRecords = currentExpertise.Count;
            Count.Text = "Количество " + filteredRecords + " из " + totalRecords;
        }
        private void PoiskPoAdress_TextChanged(object sender, TextChangedEventArgs e) => Update();
        private void FiltrKomnat_SelectionChanged(object sender, SelectionChangedEventArgs e) => Update();

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Manager.MainFrame.Navigate(new AddEditExpertise(null, _expert));
            Update();
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (_expert == null)
            {
                MessageBox.Show("Невозможно выполнить Удаление, войдите как эксперт");
            }
            else
            {
                var currentExpertise = (sender as Button).DataContext as Expertise;
                if (currentExpertise.Status == "Завершена")
                    MessageBox.Show("Невозможно выполнить удаление, данная экспертиза завершена");
                else
                {
                    if (MessageBox.Show("Вы точно хотите выполнить удаление?", "Внимание!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        try
                        {
                            currentExpertise.Status = "Отменена";
                            currentExpertise.Date_End = DateTime.Now;
                            ExpertiseEntities.GetContext().SaveChanges();
                            MessageBox.Show("Экспертиза успешно отменена", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                            Update();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }
                }
            }
            Update();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            if (_expert == null)
            {
                MessageBox.Show("Невозможно выполнить Редактирование, войдите как эксперт");
            }
            else
            {
                Manager.MainFrame.Navigate(new AddEditExpertise((sender as Button).DataContext as Expertise, _expert));
            }
            Update();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Manager.MainFrame.Navigate(new CancelExpertise(_expert));
        }


        private void ExpertiseListView_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e) => Update();

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            Update();
        }
    }
}