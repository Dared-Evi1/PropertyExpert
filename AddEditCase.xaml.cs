using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace Diplom
{
    public partial class AddEditCase : Window
    {
        private Case _currentCase;
        private bool _isEditMode;
        private string _originalNumberCase; // Для отслеживания изменений номера

        public AddEditCase(Case selectedCase)
        {
            InitializeComponent();

            if (selectedCase != null)
            {
                // Режим редактирования - создаем копию данных
                _currentCase = new Case
                {
                    Number_Case = selectedCase.Number_Case,
                    Date_Case = selectedCase.Date_Case,
                    court = selectedCase.court
                };
                _originalNumberCase = selectedCase.Number_Case; // Запоминаем оригинальный номер
                _isEditMode = true;
                this.Title = "Редактирование уголовного дела";
            }
            else
            {
                // Режим создания нового дела
                _currentCase = new Case
                {
                    Date_Case = DateTime.Now
                };
                _isEditMode = false;
                this.Title = "Новое уголовное дело";
            }

            this.DataContext = _currentCase;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Простые проверки
            if (string.IsNullOrWhiteSpace(_currentCase.Number_Case))
            {
                MessageBox.Show("Укажите номер дела", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_currentCase.Date_Case == null || _currentCase.Date_Case == DateTime.MinValue)
            {
                MessageBox.Show("Укажите дату дела", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_currentCase.Date_Case > DateTime.Now)
            {
                MessageBox.Show("Дата дела не может быть в будущем", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(_currentCase.court))
            {
                MessageBox.Show("Укажите название суда", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var context = ExpertiseEntities.GetContext();

            if (!_isEditMode)
            {
                // Проверка уникальности номера для нового дела
                var existingCase = context.Case.Find(_currentCase.Number_Case);
                if (existingCase != null)
                {
                    MessageBox.Show("Дело с таким номером уже существует", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Добавляем новое дело
                context.Case.Add(_currentCase);
            }
            else
            {
                // Для редактирования - если номер изменился, проверяем уникальность
                if (_originalNumberCase != _currentCase.Number_Case)
                {
                    var existingCase = context.Case.Find(_currentCase.Number_Case);
                    if (existingCase != null)
                    {
                        MessageBox.Show("Дело с таким номером уже существует", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }

                // Находим оригинальное дело в БД и обновляем его
                var caseToUpdate = context.Case.Find(_originalNumberCase);
                if (caseToUpdate != null)
                {
                    caseToUpdate.Number_Case = _currentCase.Number_Case;
                    caseToUpdate.Date_Case = _currentCase.Date_Case;
                    caseToUpdate.court = _currentCase.court;
                }
            }

            // Сохраняем изменения
            context.SaveChanges();

            this.DialogResult = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // Просто закрываем окно без сохранения
            this.DialogResult = false;
            this.Close();
        }
    }
}