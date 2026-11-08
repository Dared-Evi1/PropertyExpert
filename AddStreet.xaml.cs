using System;
using System.Linq;
using System.Windows;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для AddEditStreet.xaml
    /// </summary>
    public partial class AddEditStreet : Window
    {
        private Street _currentStreet;
        private bool _isEditMode;
        private int _originalStreetId;
        private string _originalStreetName;

        public AddEditStreet(Street selectedStreet = null)
        {
            InitializeComponent();

            if (selectedStreet != null)
            {
                // Режим редактирования
                _currentStreet = new Street
                {
                    Street_ID = selectedStreet.Street_ID,
                    Name_Street = selectedStreet.Name_Street
                };
                _originalStreetId = selectedStreet.Street_ID;
                _originalStreetName = selectedStreet.Name_Street;
                _isEditMode = true;
                this.Title = "Редактирование улицы";
            }
            else
            {
                // Режим создания новой улицы
                _currentStreet = new Street
                {
                    Street_ID = 0
                };
                _isEditMode = false;
                this.Title = "Новая улица";
            }

            this.DataContext = _currentStreet;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка на пустое поле
            if (string.IsNullOrWhiteSpace(_currentStreet.Name_Street))
            {
                MessageBox.Show("Введите название улицы", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var context = ExpertiseEntities.GetContext();

            // Проверка уникальности названия (без учета регистра)
            var existingStreet = context.Street.FirstOrDefault(s => s.Name_Street.ToLower() == _currentStreet.Name_Street.ToLower());

            if (existingStreet != null)
            {
                MessageBox.Show($"Улица с названием \"{_currentStreet.Name_Street}\" уже существует",
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Для новой улицы - находим последний ID и добавляем +1
            if (!_isEditMode)
            {
                // Находим максимальный существующий ID
                int maxId = 0;
                if (context.Street.Any())
                {
                    maxId = context.Street.Max(s => s.Street_ID);
                }
                // Устанавливаем новый ID = максимальный ID + 1
                _currentStreet.Street_ID = maxId + 1;
            }

            // Добавляем новую улицу
            context.Street.Add(_currentStreet);

            try
            {
                // Сохраняем изменения
                context.SaveChanges();

                this.DialogResult = true;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}