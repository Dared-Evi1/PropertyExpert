using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace Diplom
{

    public partial class AddEditExperts : Page
    {
        private Experts _currentExpert = new Experts();
        public Experts expert1;
        public bool _isUpdatingPhone = false;
        public AddEditExperts(Experts SelectedExpert)
        {
            expert1 = SelectedExpert;
            InitializeComponent();

            // Загружаем список должностей
            var positions = ExpertiseEntities.GetContext().Post.ToList();
            Post.ItemsSource = positions;
            Post.DisplayMemberPath = "Post_Name";
            Post.SelectedValuePath = "ID_Post";

            if (SelectedExpert != null)
            {
                _currentExpert = SelectedExpert;
                DataContext = _currentExpert;
                //Телефон
                PhoneTextBox.Text = FormatPhoneForDisplay(_currentExpert.Phone);
                // Устанавливаем выбранную должность для существующего эксперта
                Post.SelectedValue = _currentExpert.ID_Post;
            }
            else
            {
                _currentExpert = new Experts();
                // Установка значений по умолчанию для нового эксперта
                _currentExpert.Date_Work = DateTime.Now;
                _currentExpert.ID_Post = 2;
                DataContext = _currentExpert;
                Post.SelectedIndex = 1;
                // Устанавливаем начальную маску
                PhoneTextBox.Text = "8-";
                _currentExpert.Birthday = new DateTime(1970, 1, 1);
                _currentExpert.Status = "Работает";
            }
            if (_currentExpert.Kurs == "Пройдены")
            {
                Pro.IsChecked = true;
            }
            else
            {
                NePro.IsChecked = true;
            }
            LoadExpertImage();
        }

        private void LoadExpertImage()
        {
            try
            {
                if (!string.IsNullOrEmpty(_currentExpert.Images))
                {
                    string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                    string resFolder = System.IO.Path.Combine(projectRoot, "Res");
                    string imagePath = System.IO.Path.Combine(resFolder, _currentExpert.Images);

                    if (File.Exists(imagePath))
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        ExpertImage.Source = bitmap;
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
                            ExpertImage.Source = bitmap;
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
                        ExpertImage.Source = bitmap;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка загрузки изображения: {ex.Message}");
            }
        }

        private void ChangePicture_Click(object sender, RoutedEventArgs e)
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
                    _currentExpert.Images = fileName;

                    // Загружаем изображение для отображения
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(destinationPath, UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    ExpertImage.Source = bitmap;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе изображения: {ex.Message}");
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder errors = new StringBuilder();
            // Присваиваем значения из DatePicker
            if (Birthday.SelectedDate != null)
                _currentExpert.Birthday = Birthday.SelectedDate.Value;

            if (Date_Work.SelectedDate != null)
                _currentExpert.Date_Work = Date_Work.SelectedDate.Value;

            // Проверка обязательных полей
            if (string.IsNullOrWhiteSpace(_currentExpert.Surname))
                errors.AppendLine("Укажите Фамилию");

            if (string.IsNullOrWhiteSpace(_currentExpert.Name))
                errors.AppendLine("Укажите Имя");

            if (string.IsNullOrWhiteSpace(_currentExpert.Patronymic))
                errors.AppendLine("Укажите Отчество");

            if (Post.SelectedValue == null)
                errors.AppendLine("Выберите Должность");
            else
                _currentExpert.ID_Post = (int)Post.SelectedValue;

            // Проверка даты рождения (должна быть установлена)
            if (Birthday.SelectedDate == null)
                errors.AppendLine("Укажите Дату рождения");
            else if (Birthday.SelectedDate.Value.Year < 1900 || Birthday.SelectedDate.Value > DateTime.Today)
                errors.AppendLine("Укажите корректную Дату рождения");

            // Проверка даты приема на работу
            if (Date_Work.SelectedDate == null)
                errors.AppendLine("Укажите Дату приема на работу");
            else if (Date_Work.SelectedDate.Value.Year < 2000 || Date_Work.SelectedDate.Value > DateTime.Today)
                errors.AppendLine("Укажите корректную Дату приема на работу");
            else
            {
                _currentExpert.Date_Work = Date_Work.SelectedDate.Value;

            }


            if (string.IsNullOrWhiteSpace(_currentExpert.Email))
                errors.AppendLine("Укажите Email!");
            else
            {

                string pattern = @"^[^@\s]+@[^@\s]{3,}\.[^@\s]{2,}$";;
                const string russianLetters = "абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ";
                if (!Regex.IsMatch(_currentExpert.Email, pattern) || _currentExpert.Email.Any(russianLetters.Contains))
                    errors.AppendLine("Укажите правильный Email! (только латиница, между @ и . не менее 3 символов, домен не менее 2)");



            }


            // Проверка телефона
            if (string.IsNullOrWhiteSpace(_currentExpert.Phone))
                errors.AppendLine("Укажите номер телефона!");
            else
            {
                string clearPhone = _currentExpert.Phone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Replace("+", "");
                if (!clearPhone.All(char.IsDigit))
                    errors.AppendLine("Телефон может содержать только цифры");
                else
                {
                    if (_currentExpert.Phone.Length != 11)
                    {
                        errors.AppendLine("Номер телефона должен содержать 11 цифр");

                    }
                }
            }


            // Проверка логина
            if (string.IsNullOrWhiteSpace(_currentExpert.Login))
                errors.AppendLine("Укажите Логин");

            // Проверка пароля
            if (string.IsNullOrWhiteSpace(_currentExpert.Password))
                errors.AppendLine("Укажите Пароль");
            else
            {
                if (!Regex.IsMatch(_currentExpert.Password, @"[A-Za-zА-Яа-я]"))
                    errors.AppendLine("Пароль должен содержать хотя бы одну букву");

                if (!Regex.IsMatch(_currentExpert.Password, @"\d"))
                    errors.AppendLine("Пароль должен содержать хотя бы одну цифру");

                if (!Regex.IsMatch(_currentExpert.Password, @"[^A-Za-zА-Яа-я\d]"))
                    errors.AppendLine("Пароль должен содержать хотя бы один специальный символ");
            }

            // Проверка возраста (должен быть старше 18 на момент приема)
            if (Birthday.SelectedDate != null && Date_Work.SelectedDate != null &&
                Birthday.SelectedDate.Value.AddYears(18) > Date_Work.SelectedDate.Value)
            {
                errors.AppendLine("Эксперт должен быть старше 18 лет на момент приема на работу");
            }
            else
            {
                _currentExpert.Birthday = Birthday.SelectedDate.Value;

            }


            //Проверка курсов(пройдены ли они)
            if (Pro.IsChecked == true)
            {
                _currentExpert.Kurs = "Пройдены";
            }
            else if (NePro.IsChecked == true)
            {
                _currentExpert.Kurs = "Не Пройдены";
            }


            //Проверка образования
            if (NePro.IsChecked == true)
            {
                errors.AppendLine("Данный сотрудник не может работать экспертом");
            }

            if (Obrazov.SelectedValue == null)
                errors.AppendLine("Выберите Образование");
            else
                _currentExpert.Education = Obrazov.Text;


            // Если есть ошибки - показываем их
            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }
            try
            {
                var context = ExpertiseEntities.GetContext();

                // Сохраняем недвижимость
                if (_currentExpert.ID_Expert == 0)
                {
                    context.Experts.Add(_currentExpert);
                }
                context.SaveChanges();
                MessageBox.Show("Эксперт успешно сохранены");
                Manager.MainFrame.GoBack();


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}\n\n{ex.InnerException?.Message}");
            }
        }

        public void PhoneTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingPhone) return;
            _isUpdatingPhone = true;

            var textBox = sender as TextBox;
            if (textBox == null) return;

            // Получаем только цифры из введённого текста
            string rawInput = new string(textBox.Text.Where(char.IsDigit).ToArray());

            // Если пусто или нет цифры 8 в начале - добавляем 8
            if (string.IsNullOrEmpty(rawInput))
            {
                textBox.Text = "8-";
                textBox.CaretIndex = textBox.Text.Length;
                _isUpdatingPhone = false;
                return;
            }

            // Убеждаемся, что первая цифра - 8
            if (!rawInput.StartsWith("8"))
                rawInput = "8" + rawInput;

            // Ограничиваем длину: 1 (восьмёрка) + 10 цифр номера = 11 цифр
            if (rawInput.Length > 11)
                rawInput = rawInput.Substring(0, 11);

            // Форматируем по маске 8-xxx-xxx-xx-xx
            string formatted = "";
            if (rawInput.Length >= 1)
            {
                formatted = rawInput[0].ToString();
                if (rawInput.Length >= 2)
                    formatted += "-" + rawInput.Substring(1, Math.Min(3, rawInput.Length - 1));
                if (rawInput.Length >= 5)
                    formatted += "-" + rawInput.Substring(4, Math.Min(3, rawInput.Length - 4));
                if (rawInput.Length >= 8)
                    formatted += "-" + rawInput.Substring(7, Math.Min(2, rawInput.Length - 7));
                if (rawInput.Length >= 10)
                    formatted += "-" + rawInput.Substring(9, Math.Min(2, rawInput.Length - 9));
            }

            // Временно отключаем события, чтобы избежать зацикливания
            textBox.TextChanged -= PhoneTextBox_TextChanged;
            textBox.Text = formatted;
            textBox.TextChanged += PhoneTextBox_TextChanged;

            // Ставим курсор в конец
            textBox.CaretIndex = textBox.Text.Length;

            // Сохраняем в модель ТОЛЬКО цифры (с 8)
            _currentExpert.Phone = rawInput;

            _isUpdatingPhone = false;
        }

        public string FormatPhoneForDisplay(string rawPhone)
        {
            if (string.IsNullOrEmpty(rawPhone))
                return "8-";

            string digits = new string(rawPhone.Where(char.IsDigit).ToArray());
            if (digits.Length == 0)
                return "8-";

            if (!digits.StartsWith("8"))
                digits = "8" + digits;

            if (digits.Length > 11)
                digits = digits.Substring(0, 11);

            if (digits.Length <= 1)
                return digits;
            if (digits.Length <= 4)
                return $"{digits[0]}-{digits.Substring(1)}";
            if (digits.Length <= 7)
                return $"{digits[0]}-{digits.Substring(1, 3)}-{digits.Substring(4)}";
            if (digits.Length <= 9)
                return $"{digits[0]}-{digits.Substring(1, 3)}-{digits.Substring(4, 3)}-{digits.Substring(7)}";
            // 10-11 цифр
            return $"{digits[0]}-{digits.Substring(1, 3)}-{digits.Substring(4, 3)}-{digits.Substring(7, 2)}-{digits.Substring(9)}";
        }


    }
}