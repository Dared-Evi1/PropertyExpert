using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace Diplom
{
    public partial class AddEditAnalog : Window
    {
        private Analog_Properties _currentAnalog;
        private bool _isEditMode = false;
        private List<Street> _streets;
        public AddEditAnalog(Analog_Properties analogToEdit = null)
        {
            InitializeComponent();

            var context = ExpertiseEntities.GetContext();
            _streets = context.Street.ToList()
                .OrderBy(s => s.Name_Street)  // Добавьте эту сортировку
                .ToList();
            StreetComboBox.ItemsSource = _streets;
            //LoadSelectedStreet();
            if (analogToEdit != null)
            {
                // Режим редактирования существующего аналога
                _currentAnalog = analogToEdit;


                

                _isEditMode = true;
                Title = "Редактирование аналога";
            }
            else
            {
                // Режим создания нового аналога
                _currentAnalog = new Analog_Properties();
                _isEditMode = false;
                Title = "Добавление нового аналога";

                // Установка значений по умолчанию
                _currentAnalog.Listing_Date = DateTime.Now;
                _currentAnalog.Furniture = "Нет";
                _currentAnalog.Finishing = "Без отделки";
                _currentAnalog.Wall_Material = "Кирпич";
                _currentAnalog.Balcony = "не имеется";
                _currentAnalog.Purpose = 1; // 1 комната по умолчанию
                _currentAnalog.Price = 0; // Начальная цена
            }

            DataContext = _currentAnalog;
            LoadAnalogImage();
        }

        private void LoadAnalogImage()
        {
            try
            {
                if (!string.IsNullOrEmpty(_currentAnalog.Image))
                {
                    string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                    string resFolder = System.IO.Path.Combine(projectRoot, "Res");
                    string imagePath = System.IO.Path.Combine(resFolder, _currentAnalog.Image);

                    if (File.Exists(imagePath))
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        AnalogImage.Source = bitmap;
                    }
                    else
                    {
                        // Изображение не найдено - устанавливаем изображение по умолчанию
                        string defaultImagePath = System.IO.Path.Combine(resFolder, "PictureRealties.jpeg");
                        if (File.Exists(defaultImagePath))
                        {
                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.UriSource = new Uri(defaultImagePath, UriKind.Absolute);
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.EndInit();
                            AnalogImage.Source = bitmap;
                        }
                    }
                }
                else
                {
                    // Нет изображения - устанавливаем изображение по умолчанию
                    string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                    string resFolder = System.IO.Path.Combine(projectRoot, "Res");
                    string defaultImagePath = System.IO.Path.Combine(resFolder, "PictureRealties.jpeg");
                    if (File.Exists(defaultImagePath))
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(defaultImagePath, UriKind.Absolute);
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        AnalogImage.Source = bitmap;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка загрузки изображения: {ex.Message}");
            }
        }

        private void LoadImage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog myOpenFileDialog = new OpenFileDialog();
                myOpenFileDialog.Filter = "Image files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All files (*.*)|*.*";

                string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                string resFolder = System.IO.Path.Combine(projectRoot, "Res");

                if (!Directory.Exists(resFolder))
                {
                    Directory.CreateDirectory(resFolder);
                }

                myOpenFileDialog.InitialDirectory = resFolder;

                if (myOpenFileDialog.ShowDialog() == true)
                {
                    string selectedFilePath = myOpenFileDialog.FileName;

                    // Копируем файл в папку Res, если его там нет
                    string fileName = System.IO.Path.GetFileName(selectedFilePath);
                    string destinationPath = System.IO.Path.Combine(resFolder, fileName);

                    if (!File.Exists(destinationPath))
                    {
                        File.Copy(selectedFilePath, destinationPath, true);
                    }

                    // Сохраняем только имя файла
                    _currentAnalog.Image = fileName;

                    // Загружаем изображение для отображения
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(destinationPath, UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    AnalogImage.Source = bitmap;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе изображения: {ex.Message}");
            }
        }



        private void DistrictComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DistrictComboBox.SelectedItem is TextBlock selectedTextBlock)
            {
                _currentAnalog.District = selectedTextBlock.Text;
            }
        }
        private void AddStreet_Click(object sender, RoutedEventArgs e)
        {

            var StreetWindow = new AddEditStreet(null);
            StreetWindow.Owner = Window.GetWindow(this);

            if (StreetWindow.ShowDialog() == true)
            {
                RefreshStreetList();
                if (StreetWindow.DataContext is Street newStreet && newStreet != null)
                {
                    StreetComboBox.SelectedValue = newStreet.Street_ID;
                    _currentAnalog.ID_Street = newStreet.Street_ID;
                }
            }
        }


        private void RefreshStreetList()
        {
            var context = ExpertiseEntities.GetContext();
            var street = context.Street.ToList()
                .OrderBy(s => s.Name_Street)  // Добавьте сортировку
                .ToList();

            var currentSelection = StreetComboBox.SelectedValue;

            StreetComboBox.ItemsSource = street;
            StreetComboBox.DisplayMemberPath = "Name_Street";
            StreetComboBox.SelectedValuePath = "Street_ID";

            if (currentSelection != null)
            {
                StreetComboBox.SelectedValue = currentSelection;
            }
        }

        /*
        private void LoadSelectedStreet()
        {
            if (_currentAnalog != null)
            {
                // Устанавливаем выбранное значение в ComboBox
                StreetComboBox.SelectedValue = _currentAnalog.ID_Street.ToString();
            }
        }
        */
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder errors = new StringBuilder();

            // Проверка обязательных полей


            if (string.IsNullOrWhiteSpace(_currentAnalog.District))
                errors.AppendLine("Укажите Район");

            if (string.IsNullOrWhiteSpace(_currentAnalog.House))
                errors.AppendLine("Укажите Дом");

            if (_currentAnalog.Area <= 0)
                errors.AppendLine("Укажите корректную Площадь (больше 0)");

            if (_currentAnalog.Floor <= 0 || _currentAnalog.Floor > _currentAnalog.Total_Floors)
                errors.AppendLine("Укажите корректный Этаж");

            if (_currentAnalog.Total_Floors <= 0)
                errors.AppendLine("Укажите корректное общее количество этажей");

            if (_currentAnalog.Purpose <= 0)
                errors.AppendLine("Укажите корректное количество комнат");

            if (StreetComboBox.SelectedItem == null)
                errors.AppendLine("Укажите Улицу");
            else
                _currentAnalog.ID_Street = (int)StreetComboBox.SelectedValue;


            // Проверка даты размещения
            if (_currentAnalog.Listing_Date.Year < 2010 || _currentAnalog.Listing_Date.Date > DateTime.Today)
                errors.AppendLine("Укажите корректную дату размещения объявления (2010-сегодня)");

            // Проверка стоимости (должна быть больше 0)
            if (_currentAnalog.Price <= 0)
                errors.AppendLine("Укажите корректную стоимость аналога (больше 0)");

            // Проверка ComboBox
            if (string.IsNullOrWhiteSpace(_currentAnalog.Balcony))
                errors.AppendLine("Укажите наличие Балкона");

            if (string.IsNullOrWhiteSpace(_currentAnalog.Furniture))
                errors.AppendLine("Укажите наличие Мебели");

            if (string.IsNullOrWhiteSpace(_currentAnalog.Wall_Material))
                errors.AppendLine("Укажите Материал стен");

            if (string.IsNullOrWhiteSpace(_currentAnalog.Finishing))
                errors.AppendLine("Укажите Отделку");

            // Если есть ошибки - показываем и выходим
            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString(), "Ошибки валидации",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                var context = ExpertiseEntities.GetContext();

                if (_isEditMode)
                {
                    // Режим редактирования
                    // Находим существующий аналог в базе
                    var existingAnalog = context.Analog_Properties.Find(_currentAnalog.ID_Analog_Property);

                    if (existingAnalog != null)
                    {
                        // Копируем изменения из _currentAnalog в existingAnalog
                        context.Entry(existingAnalog).CurrentValues.SetValues(_currentAnalog);

                        // Обновляем навигационные свойства если нужно

                        existingAnalog.District = _currentAnalog.District;
                        existingAnalog.Neighborhood = _currentAnalog.Neighborhood;
                        existingAnalog.House = _currentAnalog.House;
                        existingAnalog.Area = _currentAnalog.Area;
                        existingAnalog.Floor = _currentAnalog.Floor;
                        existingAnalog.Total_Floors = _currentAnalog.Total_Floors;
                        existingAnalog.Balcony = _currentAnalog.Balcony;
                        existingAnalog.Price = _currentAnalog.Price;
                        existingAnalog.Furniture = _currentAnalog.Furniture;
                        existingAnalog.Finishing = _currentAnalog.Finishing;
                        existingAnalog.Wall_Material = _currentAnalog.Wall_Material;
                        existingAnalog.Purpose = _currentAnalog.Purpose;
                        existingAnalog.Listing_Date = _currentAnalog.Listing_Date;
                        existingAnalog.Image = _currentAnalog.Image;

                    }
                    else
                    {
                        MessageBox.Show("Не удалось найти аналог для редактирования", "Ошибка",
                                        MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
                else
                {
                    context.Analog_Properties.Add(_currentAnalog);
                }

                // Сохраняем изменения (один раз!)
                context.SaveChanges();

                string message = _isEditMode
                    ? "Аналог успешно обновлен!"
                    : "Новый аналог успешно добавлен!";

                MessageBox.Show(message, "Успешно",
                                MessageBoxButton.OK, MessageBoxImage.Information);

                this.DialogResult = true;
                this.Close();
            }
            catch (Exception ex)
            {
                string errorMessage = $"Ошибка при сохранении аналога: {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMessage += $"\nВнутренняя ошибка: {ex.InnerException.Message}";

                    // Для отладки: показываем полный стек вызовов
                    if (ex.InnerException.InnerException != null)
                    {
                        errorMessage += $"\nДетали: {ex.InnerException.InnerException.Message}";
                    }
                }
                MessageBox.Show(errorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}