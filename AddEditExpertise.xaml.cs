using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Entity;
using System.Diagnostics;
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
using Path = System.IO.Path;
using Word = Microsoft.Office.Interop.Word;

namespace Diplom
{
    public partial class AddEditExpertise : Page
    {
        private Expertise _currentExpertise = new Expertise();
        private Realties _currentRealty = new Realties();
        private ObservableCollection<Analog_Properties> _availableAnalogs;
        private ObservableCollection<Analog_Expertise> _selectedAnalogs;
        private List<Street> _streets;
        public Experts expert1;

        // Вспомогательный класс для отображения аналогов в ListBox
        public class AnalogWrapper
        {
            public Analog_Properties Analog { get; set; }
            public string Address => Analog?.FullAddress_Analog ?? "Адрес не указан";
            public string Area => Analog?.Area_Analog ?? "0 м²";
            public string Floor => Analog?.Floor_Analog ?? "0/0";
            public string Price => Analog?.Price.HasValue == true ? $"{Analog.Price:N0} ₽" : "Цена не указана";
            public string Rooms => Analog?.Purpose.ToString() + "-комн.";
        }

        public AddEditExpertise(Expertise selectedExpertise, Experts selectedExpert)
        {
            InitializeComponent();

            var experts = ExpertiseEntities.GetContext().Experts.ToList().Where(p=> p.Status == "Работает");
            ExpertComboBox.ItemsSource = experts;
            ExpertComboBox.DisplayMemberPath = "FIO";
            ExpertComboBox.SelectedValuePath = "ID_Expert";

            // ЗАГРУЖАЕМ СУДЕБНЫЕ ДЕЛА
            var cases = ExpertiseEntities.GetContext().Case.ToList()
                .OrderBy(c => c.Number_Case)
                .ToList();
            CaseComboBox.ItemsSource = cases;
            CaseComboBox.DisplayMemberPath = "Number_Case";
            CaseComboBox.SelectedValuePath = "Number_Case";

            var context = ExpertiseEntities.GetContext();
            _streets = context.Street.ToList()
                .OrderBy(s => s.Name_Street)
                .ToList();
            StreetComboBox.ItemsSource = _streets;

            _availableAnalogs = new ObservableCollection<Analog_Properties>();
            _selectedAnalogs = new ObservableCollection<Analog_Expertise>();

            if (selectedExpertise != null)
            {
                _currentExpertise = selectedExpertise;
                _currentRealty = selectedExpertise.Realties ?? new Realties();

                if (_currentExpertise.ID_Expert > 0)
                    ExpertComboBox.SelectedValue = _currentExpertise.ID_Expert;

                if (_currentExpertise.ID_Case != null)
                    CaseComboBox.SelectedValue = _currentExpertise.ID_Case;

                LoadSelectedAnalogs();
            }
            else
            {
                _currentExpertise = new Expertise();
                _currentRealty = new Realties();

                _currentExpertise.Date_Start = DateTime.Now;
                _currentExpertise.Status = "В процессе";
                _currentExpertise.Cost = 0;
                _currentExpertise.ID_Case = "2-1/2023";
                _currentRealty.District = "Кировский";
                _currentRealty.ID_Street = 37;
                _currentRealty.Listing_Date = DateTime.Now;
                _currentRealty.Furniture = "Нет";
                _currentRealty.Finishing = "Без отделки";
                _currentRealty.Wall_Material = "Кирпич";
                _currentRealty.Balcony = "не имеется";
                _currentRealty.Purpose = 1;

                if (selectedExpert != null)
                {
                    _currentExpertise.ID_Expert = selectedExpert.ID_Expert;
                    ExpertComboBox.SelectedValue = selectedExpert.ID_Expert;
                }
            }

            LoadAvailableAnalogs();
            DataContext = new
            {
                Realties = _currentRealty,
                Expertise = _currentExpertise,
                AvailableAnalogs = _availableAnalogs,
                SelectedAnalogs = _selectedAnalogs
            };

            LoadRealtyImage();
        }

        private void LoadRealtyImage()
        {
            try
            {
                if (!string.IsNullOrEmpty(_currentRealty.Image))
                {
                    string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                    string resFolder = System.IO.Path.Combine(projectRoot, "Res");
                    string imagePath = System.IO.Path.Combine(resFolder, _currentRealty.Image);

                    if (File.Exists(imagePath))
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        RealtyImage.Source = bitmap;
                    }
                    else
                    {
                        string defaultImagePath = System.IO.Path.Combine(resFolder, "PictureRealties.jpeg");
                        if (File.Exists(defaultImagePath))
                        {
                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.UriSource = new Uri(defaultImagePath, UriKind.Absolute);
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.EndInit();
                            RealtyImage.Source = bitmap;
                        }
                    }
                }
                else
                {
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
                        RealtyImage.Source = bitmap;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка загрузки изображения: {ex.Message}");
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
                    _currentRealty.ID_Street = newStreet.Street_ID;
                }
            }
        }

        private void RefreshStreetList()
        {
            var context = ExpertiseEntities.GetContext();
            var street = context.Street.ToList()
                .OrderBy(s => s.Name_Street)
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

        private void PurposeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var binding = PurposeTextBox.GetBindingExpression(TextBox.TextProperty);
            binding?.UpdateSource();

            Dispatcher.BeginInvoke(new Action(() =>
            {
                LoadAvailableAnalogs();
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void CaseComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CaseComboBox.SelectedItem != null)
            {
                var selectedCase = CaseComboBox.SelectedItem as Case;
                if (selectedCase != null)
                {
                    _currentExpertise.ID_Case = selectedCase.Number_Case;
                }
            }
        }

        private void DistrictComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DistrictComboBox.SelectedItem is TextBlock selectedTextBlock)
            {
                _currentRealty.District = selectedTextBlock.Text;
                LoadAvailableAnalogs();
            }
        }

        private void LoadAvailableAnalogs()
        {
            try
            {
                var context = ExpertiseEntities.GetContext();
                var allAnalogs = context.Analog_Properties.ToList();

                if (_currentExpertise.ID_Expertise > 0)
                {
                    var selectedAnalogIds = _selectedAnalogs.Select(a => a.ID_Vnes_Analog_Property).ToList();
                    allAnalogs = allAnalogs.Where(a => !selectedAnalogIds.Contains(a.ID_Analog_Property)).ToList();
                }

                int currentPurpose = _currentRealty.Purpose;
                string currentDistrict = _currentRealty.District;

                IEnumerable<Analog_Properties> filteredAnalogs = allAnalogs;

                if (!string.IsNullOrWhiteSpace(currentDistrict))
                {
                    filteredAnalogs = filteredAnalogs.Where(a =>
                        a.District != null &&
                        a.District.Trim().Equals(currentDistrict.Trim(), StringComparison.OrdinalIgnoreCase));
                }

                if (currentPurpose > 0)
                {
                    filteredAnalogs = filteredAnalogs.Where(a => a.Purpose == currentPurpose);
                }

                var result = filteredAnalogs.ToList();

                _availableAnalogs.Clear();
                foreach (var analog in result)
                {
                    _availableAnalogs.Add(analog);
                }

                AvailableAnalogsList.ItemsSource = _availableAnalogs.Select(a => new AnalogWrapper { Analog = a }).ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки аналогов: {ex.Message}");
            }
        }

        private void LoadSelectedAnalogs()
        {
            try
            {
                var context = ExpertiseEntities.GetContext();
                var selectedAnalogs = context.Analog_Expertise
                    .Where(ae => ae.ID_Vnes_Expertise == _currentExpertise.ID_Expertise)
                    .Include(ae => ae.Analog_Properties)
                    .ToList();

                _selectedAnalogs.Clear();
                foreach (var analog in selectedAnalogs)
                {
                    _selectedAnalogs.Add(analog);
                }

                SelectedAnalogsList.ItemsSource = _selectedAnalogs
                    .Select(ae => new AnalogWrapper
                    {
                        Analog = ae.Analog_Properties
                    })
                    .ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки выбранных аналогов: {ex.Message}");
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
                    string fileName = System.IO.Path.GetFileName(selectedFilePath);
                    string destinationPath = System.IO.Path.Combine(resFolder, fileName);

                    if (!File.Exists(destinationPath))
                    {
                        File.Copy(selectedFilePath, destinationPath, true);
                    }

                    _currentRealty.Image = fileName;

                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(destinationPath, UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    RealtyImage.Source = bitmap;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе изображения: {ex.Message}");
            }
        }

        private void AddAnalogButton_Click(object sender, RoutedEventArgs e)
        {
            if (AvailableAnalogsList.SelectedItem is AnalogWrapper selectedWrapper && selectedWrapper.Analog != null)
            {
                var selectedAnalog = selectedWrapper.Analog;
                var context = ExpertiseEntities.GetContext();

                var fullAnalog = context.Analog_Properties.Find(selectedAnalog.ID_Analog_Property);

                if (fullAnalog == null)
                {
                    MessageBox.Show("Не удалось загрузить данные аналога");
                    return;
                }

                if (!_selectedAnalogs.Any(ae => ae.ID_Vnes_Analog_Property == fullAnalog.ID_Analog_Property))
                {
                    var analogExpertise = new Analog_Expertise
                    {
                        ID_Vnes_Analog_Property = fullAnalog.ID_Analog_Property,
                        ID_Vnes_Expertise = _currentExpertise.ID_Expertise,
                        check = 1,
                        Analog_Properties = fullAnalog
                    };

                    _selectedAnalogs.Add(analogExpertise);
                    _availableAnalogs.Remove(selectedAnalog);

                    AvailableAnalogsList.ItemsSource = _availableAnalogs
                        .Select(a => new AnalogWrapper { Analog = a })
                        .ToList();

                    SelectedAnalogsList.ItemsSource = _selectedAnalogs
                        .Select(ae => new AnalogWrapper
                        {
                            Analog = ae.Analog_Properties ??
                                     context.Analog_Properties.Find(ae.ID_Vnes_Analog_Property)
                        })
                        .ToList();
                }
                else
                {
                    MessageBox.Show("Этот аналог уже добавлен в экспертизу");
                }
            }
            else
            {
                MessageBox.Show("Выберите аналог для добавления");
            }
        }

        private void RemoveAnalogButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedAnalogsList.SelectedItem is AnalogWrapper selectedWrapper && selectedWrapper.Analog != null)
            {
                var selectedAnalog = selectedWrapper.Analog;
                var analogToRemove = _selectedAnalogs.FirstOrDefault(ae =>
                    ae.ID_Vnes_Analog_Property == selectedAnalog.ID_Analog_Property);

                if (analogToRemove != null)
                {
                    _selectedAnalogs.Remove(analogToRemove);
                    _availableAnalogs.Add(selectedAnalog);

                    AvailableAnalogsList.ItemsSource = _availableAnalogs.Select(a => new AnalogWrapper { Analog = a });
                    SelectedAnalogsList.ItemsSource = _selectedAnalogs.Select(ae => new AnalogWrapper { Analog = ae.Analog_Properties });
                }
            }
            else
            {
                MessageBox.Show("Выберите аналог для удаления");
            }
        }

        private void AddAnalog_Click(object sender, RoutedEventArgs e)
        {
            var editWindow = new AddEditAnalog(null);
            editWindow.Owner = Window.GetWindow(this);
            if (editWindow.ShowDialog() == true)
            {
                LoadAvailableAnalogs();
            }
        }

        private void AvailableAnalogsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            AddAnalogButton.IsEnabled = AvailableAnalogsList.SelectedItem != null;
            bool canEdit = AvailableAnalogsList.SelectedItem != null ||
                           SelectedAnalogsList.SelectedItem != null;
            EditAnalog.IsEnabled = canEdit;
            AddAnalog.IsEnabled = true;
        }

        private void SelectedAnalogsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RemoveAnalogButton.IsEnabled = SelectedAnalogsList.SelectedItem != null;
            bool canEdit = AvailableAnalogsList.SelectedItem != null ||
                           SelectedAnalogsList.SelectedItem != null;
            EditAnalog.IsEnabled = canEdit;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrWhiteSpace(_currentRealty.District))
                errors.AppendLine("Укажите Район");

            if (StreetComboBox.SelectedItem == null)
                errors.AppendLine("Укажите Улицу");
            else
                _currentRealty.ID_Street = (int)StreetComboBox.SelectedValue;
            if (string.IsNullOrWhiteSpace(_currentRealty.House))
                errors.AppendLine("Укажите Дом");

            if (string.IsNullOrWhiteSpace(_currentRealty.Apartment_number))
                errors.AppendLine("Укажите номер квартиры");

            if (_currentRealty.Area <= 0)
                errors.AppendLine("Укажите корректную Площадь (больше 0)");

            if (_currentRealty.Floor <= 0 || _currentRealty.Floor > _currentRealty.Total_Floors)
                errors.AppendLine("Укажите корректный Этаж");

            if (_currentRealty.Total_Floors <= 0)
                errors.AppendLine("Укажите корректное общее количество этажей");

            if (_currentRealty.Purpose <= 0)
                errors.AppendLine("Укажите корректное количество комнат");

            if (ListingDate.SelectedDate == null || ListingDate.SelectedDate.Value.Year < 2010 || ListingDate.SelectedDate.Value.Date > DateTime.Today)
                errors.AppendLine("Укажите корректную дату размещения объявления");
            else
                _currentRealty.Listing_Date = ListingDate.SelectedDate.Value;

            if (ExpertComboBox.SelectedValue == null)
                errors.AppendLine("Выберите ответственного эксперта");

            if (BalconyComboBox.SelectedValue == null)
                errors.AppendLine("Укажите наличие Балкона");

            if (FurnitureComboBox.SelectedValue == null)
                errors.AppendLine("Укажите наличие Мебели");

            if (WallMaterialComboBox.SelectedValue == null)
                errors.AppendLine("Укажите Материал стен");

            if (FinishingComboBox.SelectedValue == null)
                errors.AppendLine("Укажите Отделку");

            if (_currentExpertise.Cost <= 0)
                errors.AppendLine("Укажите корректную стоимость экспертизы");

            if (StatusComboBox.SelectedValue == null)
                errors.AppendLine("Укажите корректный статус");

            if (DateStart.SelectedDate == null || DateStart.SelectedDate.Value.Year < 2010 || DateStart.SelectedDate.Value.Date > DateTime.Today)
                errors.AppendLine("Укажите корректную дату начала экспертизы");
            else
                _currentExpertise.Date_Start = DateStart.SelectedDate.Value;

            if (DateEnd.SelectedDate.HasValue)
            {
                if (DateEnd.SelectedDate.Value.Year < 2010 || DateEnd.SelectedDate.Value.Date < DateStart.SelectedDate.Value.Date)
                    errors.AppendLine("Укажите корректную дату окончания экспертизы");
                else
                    _currentExpertise.Date_End = DateEnd.SelectedDate.Value;
            }

            if (_selectedAnalogs.Count != 4)
            {
                errors.AppendLine($"Должно быть выбрано ровно 4 аналога. Сейчас выбрано: {_selectedAnalogs.Count}");
            }

            if (_selectedAnalogs.Count > 0)
            {
                var context = ExpertiseEntities.GetContext();
                var invalidAnalogs = new List<string>();

                foreach (var analogExpertise in _selectedAnalogs)
                {
                    var analog = analogExpertise.Analog_Properties;
                    if (analog == null && analogExpertise.ID_Vnes_Analog_Property > 0)
                    {
                        analog = context.Analog_Properties.Find(analogExpertise.ID_Vnes_Analog_Property);
                    }

                    if (analog != null && analog.Purpose != _currentRealty.Purpose)
                    {
                        string address = analog.FullAddress_Analog;
                        invalidAnalogs.Add($"Аналог по адресу {address}: {analog.Purpose} комнат (должно быть {_currentRealty.Purpose})");
                    }
                }

                if (invalidAnalogs.Count > 0)
                {
                    errors.AppendLine("Выбранные аналоги имеют разное количество комнат:");
                    foreach (var invalidAnalog in invalidAnalogs)
                    {
                        errors.AppendLine($"- {invalidAnalog}");
                    }
                }
            }

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString(), "Ошибки заполнения", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                var context = ExpertiseEntities.GetContext();

                if (_currentRealty.ID_Realty == 0)
                {
                    context.Realties.Add(_currentRealty);
                    context.SaveChanges();
                }

                if (_currentExpertise.ID_Expertise == 0)
                {
                    _currentExpertise.ID_Realty = _currentRealty.ID_Realty;
                    context.Expertise.Add(_currentExpertise);
                    context.SaveChanges();
                }

                var oldAnalogExpertises = context.Analog_Expertise
                    .Where(ae => ae.ID_Vnes_Expertise == _currentExpertise.ID_Expertise)
                    .ToList();

                foreach (var oldAnalog in oldAnalogExpertises)
                {
                    context.Analog_Expertise.Remove(oldAnalog);
                }

                foreach (var selectedAnalog in _selectedAnalogs)
                {
                    var newAnalogExpertise = new Analog_Expertise
                    {
                        ID_Vnes_Expertise = _currentExpertise.ID_Expertise,
                        ID_Vnes_Analog_Property = selectedAnalog.ID_Vnes_Analog_Property
                    };
                    context.Analog_Expertise.Add(newAnalogExpertise);
                }

                decimal? calculatedPrice = CalculateAndSavePropertyValue();
                if (calculatedPrice.HasValue)
                {
                    _currentRealty.Price = calculatedPrice.Value;
                    context.Entry(_currentRealty).State = EntityState.Modified;
                }

                if (DateEnd.SelectedDate.HasValue)
                {
                    _currentExpertise.Status = "Завершена";
                }
                else
                {
                    StatusComboBox.Text = "В процессе";
                    _currentExpertise.Status = "В процессе";
                }

                context.SaveChanges();

                MessageBox.Show("Экспертиза и аналоги успешно сохранены", "Успех",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                if (calculatedPrice.HasValue && _currentExpertise.ID_Expertise > 0)
                {
                    try
                    {
                        var savedExpertise = context.Expertise
                            .Include(c => c.Realties)
                            .Include(c => c.Experts)
                            .Include(c => c.Case)
                            .FirstOrDefault(c => c.ID_Expertise == _currentExpertise.ID_Expertise);

                        if (savedExpertise != null)
                        {
                            ExportToExcel(savedExpertise);
                            ExportToWord(savedExpertise);
                            MessageBox.Show("Данные успешно экспортированы в Excel и Word!",
                                "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при экспорте: {ex.Message}",
                            "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }

                expert1 = ExpertiseEntities.GetContext().Experts.ToList().Find(p => p.ID_Expert == _currentExpertise.ID_Expert);
                Manager.MainFrame.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}\n\n{ex.InnerException?.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private decimal? CalculateAndSavePropertyValue()
        {
            try
            {
                if (_selectedAnalogs.Count != 4)
                {
                    MessageBox.Show("Для расчета стоимости необходимо выбрать ровно 4 аналога", "Внимание",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return null;
                }

                var context = ExpertiseEntities.GetContext();

                var analogs = new List<Analog_Properties>();
                foreach (var analogExpertise in _selectedAnalogs)
                {
                    var analog = context.Analog_Properties
                        .FirstOrDefault(a => a.ID_Analog_Property == analogExpertise.ID_Vnes_Analog_Property);
                    if (analog != null)
                    {
                        analogs.Add(analog);
                    }
                }

                if (analogs.Count != 4)
                {
                    MessageBox.Show("Не удалось загрузить данные всех аналогов", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return null;
                }

                decimal researchArea = _currentRealty.Area;
                int researchRooms = _currentRealty.Purpose;
                string researchFurniture = _currentRealty.Furniture;
                string researchFinishing = _currentRealty.Finishing ?? "Без отделки";
                string researchBalcony = _currentRealty.Balcony ?? "не имеется";
                string researchWallMaterial = _currentRealty.Wall_Material ?? "Кирпич";
                int researchFloor = _currentRealty.Floor;
                int researchTotalFloors = _currentRealty.Total_Floors;

                decimal marketDiscount = 0.04m;

                var analogFinalPrices = new List<decimal>();
                var analogGrossCorrections = new List<decimal>();

                for (int i = 0; i < 4; i++)
                {
                    var analog = analogs[i];

                    decimal analogPricePerSqm = (decimal)(analog.Price / analog.Area);
                    decimal currentPrice = analogPricePerSqm;
                    decimal totalGrossCorrection = 0;

                    decimal marketCorrection = -marketDiscount;
                    currentPrice = currentPrice * (1 + marketCorrection);
                    totalGrossCorrection += Math.Abs(marketCorrection);

                    decimal wallMaterialCorrection = (decimal)GetWallMaterialCorrection(researchWallMaterial, analog.Wall_Material);
                    currentPrice = currentPrice * (1 + wallMaterialCorrection);
                    totalGrossCorrection += Math.Abs(wallMaterialCorrection);

                    decimal areaCorrection = (decimal)CalculateAreaCorrection((double)researchArea, (double)analog.Area);
                    currentPrice = currentPrice * (1 + areaCorrection);
                    totalGrossCorrection += Math.Abs(areaCorrection);

                    decimal floorCorrection = (decimal)CalculateFloorCorrection(researchFloor, analog.Floor, researchTotalFloors, analog.Total_Floors);
                    currentPrice = currentPrice * (1 + floorCorrection);
                    totalGrossCorrection += Math.Abs(floorCorrection);

                    decimal finishingCorrection = (decimal)GetFinishingCorrection(researchFinishing, analog.Finishing);
                    currentPrice = currentPrice * (1 + finishingCorrection);
                    totalGrossCorrection += Math.Abs(finishingCorrection);

                    decimal furnitureCorrection = (decimal)GetFurnitureCorrection(researchFurniture, analog.Furniture);
                    currentPrice = currentPrice * (1 + furnitureCorrection);
                    totalGrossCorrection += Math.Abs(furnitureCorrection);

                    decimal balconyCorrection = (decimal)GetBalconyCorrection(researchBalcony, analog.Balcony);
                    currentPrice = currentPrice * (1 + balconyCorrection);
                    totalGrossCorrection += Math.Abs(balconyCorrection);

                    analogFinalPrices.Add(currentPrice);
                    analogGrossCorrections.Add(totalGrossCorrection);
                }

                var inverseCorrections = new List<decimal>();
                foreach (var correction in analogGrossCorrections)
                {
                    inverseCorrections.Add(1 / correction);
                }

                decimal sumInverse = inverseCorrections.Sum();
                var weights = new List<decimal>();
                foreach (var invCorr in inverseCorrections)
                {
                    weights.Add(invCorr / sumInverse);
                }

                decimal weightedPricePerSqm = 0;
                for (int i = 0; i < 4; i++)
                {
                    weightedPricePerSqm += analogFinalPrices[i] * weights[i];
                }

                decimal finalPropertyValue = weightedPricePerSqm * researchArea;

                return finalPropertyValue;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при расчете стоимости: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private void EditAnalog_Click(object sender, RoutedEventArgs e)
        {
            Analog_Properties analogToEdit = null;

            if (SelectedAnalogsList.SelectedItem is AnalogWrapper selectedWrapper && selectedWrapper.Analog != null)
            {
                analogToEdit = selectedWrapper.Analog;
            }
            else if (AvailableAnalogsList.SelectedItem is AnalogWrapper availableWrapper && availableWrapper.Analog != null)
            {
                analogToEdit = availableWrapper.Analog;
            }

            if (analogToEdit != null)
            {
                var editWindow = new AddEditAnalog(analogToEdit);
                editWindow.Owner = Window.GetWindow(this);

                if (editWindow.ShowDialog() == true)
                {
                    LoadAvailableAnalogs();
                    LoadSelectedAnalogs();
                    MessageBox.Show("Аналог успешно отредактирован", "Успех",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            else
            {
                MessageBox.Show("Выберите аналог для редактирования (из доступных или выбранных)",
                                "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void AvailableAnalogsList_GotFocus(object sender, RoutedEventArgs e)
        {
            SelectedAnalogsList.SelectedItem = null;
        }

        private void SelectedAnalogsList_GotFocus(object sender, RoutedEventArgs e)
        {
            AvailableAnalogsList.SelectedItem = null;
        }

        // ==================== МЕТОДЫ ЭКСПОРТА ====================

        private void ExportToExcel(Expertise selectedExpertise)
        {
            string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string templatePath = System.IO.Path.Combine(projectRoot, "Export", "Шаблон.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Шаблон Excel не найден по пути: {templatePath}");
            }

            string fileName = GenerateFileName(selectedExpertise.Realties) + ".xlsx";
            string exportFolder = System.IO.Path.Combine(projectRoot, "Export");

            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
            }

            string outputPath = System.IO.Path.Combine(exportFolder, fileName);

            int fileNumber = 1;
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(fileName);
            string extension = System.IO.Path.GetExtension(fileName);

            while (File.Exists(outputPath))
            {
                string newFileName = $"{fileNameWithoutExtension}_{fileNumber}{extension}";
                outputPath = System.IO.Path.Combine(exportFolder, newFileName);
                fileNumber++;
            }

            File.Copy(templatePath, outputPath, true);
            FillExcelDocument(outputPath, selectedExpertise);
        }

        private void FillExcelDocument(string filePath, Expertise expertise)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                var context = ExpertiseEntities.GetContext();
                var analogs = context.Analog_Expertise
                    .Where(ae => ae.ID_Vnes_Expertise == expertise.ID_Expertise)
                    .Select(ae => ae.Analog_Properties)
                    .Take(4)
                    .ToList();

                while (analogs.Count < 4)
                {
                    analogs.Add(null);
                }

                var realty = expertise.Realties;
                if (realty == null) return;

                // Получаем данные для расчета корректировок
                decimal researchArea = realty.Area;
                string researchFinishing = realty.Finishing ?? "Без отделки";
                string researchBalcony = realty.Balcony ?? "не имеется";
                string researchWallMaterial = realty.Wall_Material ?? "Кирпич";
                string researchFurniture = realty.Furniture ?? "Нет";
                int researchFloor = realty.Floor;
                int researchTotalFloors = realty.Total_Floors;
                decimal marketDiscount = 0.04m;

                // Заполнение данных в шаблон
                // Строка 2: Местоположение
                worksheet.Cells[2, 3] = GetFullAddress(expertise.Realties);
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[2, 4 + i] = GetFullAddress(analogs[i]);

                // Строка 3: Дата предложения
                worksheet.Cells[3, 3] = FormatDate(expertise.Realties?.Listing_Date);
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[3, 4 + i] = FormatDate(analogs[i]?.Listing_Date);

                // Строка 4: Цена предложения
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[4, 4 + i] = analogs[i]?.Price ?? 0;

                // Строка 5: Площадь объекта
                worksheet.Cells[5, 3] = researchArea;
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[5, 4 + i] = analogs[i]?.Area ?? 0;

                // Строка 6: Цена предложения за кв.м. (ФОРМУЛА)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[6, 4 + i].Formula = $"=IF({col}5>0, {col}4/{col}5, 0)";
                }

                // Строка 7: Предполагаемые имущественные права
                for (int i = 3; i <= 7; i++)
                    worksheet.Cells[7, i] = "собственность";

                // Строка 8: Корректировка на права
                for (int i = 3; i <= 7; i++)
                    worksheet.Cells[8, i] = 0;

                // Строка 9: Скорректированная стоимость (после прав)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[9, 4 + i].Formula = $"={col}6";
                }

                // Строка 10: Условия финансирования
                for (int i = 3; i <= 7; i++)
                    worksheet.Cells[10, i] = "обычные";

                // Строка 11: Корректировка на финансирование
                for (int i = 3; i <= 7; i++)
                    worksheet.Cells[11, i] = 0;

                // Строка 12: Скорректированная стоимость (после финансирования)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[12, 4 + i].Formula = $"={col}9";
                }

                // Строка 13: Условия рынка
                worksheet.Cells[13, 3] = "коммерческие";
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[13, 4 + i] = "предложение";

                // Строка 14: Корректировка на торг
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[14, 4 + i] = -marketDiscount;

                // Строка 15: Скорректированная стоимость (после торга)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[15, 4 + i].Formula = $"={col}12*(1+{col}14)";
                }

                // Строка 16: Назначение
                worksheet.Cells[16, 3] = $"квартира, {expertise.Realties?.Purpose}-к";
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[16, 4 + i] = analogs[i] == null ? "нет данных" : $"квартира, {analogs[i].Purpose}-к";

                // Строка 17: Корректировка на назначение
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[17, 4 + i] = 0;

                // Строка 18: Скорректированная стоимость (после назначения)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[18, 4 + i].Formula = $"={col}15";
                }

                // Строка 19: Материал стен
                worksheet.Cells[19, 3] = GetWallMaterialForExcel(expertise.Realties?.Wall_Material);
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[19, 4 + i] = GetWallMaterialForExcel(analogs[i]?.Wall_Material);

                // Строка 20: Корректировка на материал стен (ФОРМУЛА)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[20, 4 + i].Formula =
                        $"=IF($C$19=\"кирпич\", " +
                        $"IF(${col}$19=\"панель\", 0.05, " +
                        $"IF(${col}$19=\"дерево\", 0.08, 0)), " +
                        $"IF($C$19=\"панель\", " +
                        $"IF(${col}$19=\"кирпич\", -0.05, " +
                        $"IF(${col}$19=\"дерево\", 0.03, 0)), " +
                        $"IF($C$19=\"дерево\", " +
                        $"IF(${col}$19=\"кирпич\", -0.08, " +
                        $"IF(${col}$19=\"панель\", -0.05, 0)), 0)))";
                }

                // Строка 21: Скорректированная стоимость (после материала стен)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[21, 4 + i].Formula = $"={col}18*(1+{col}20)";
                }

                // Строка 22: Площадь
                worksheet.Cells[22, 3] = researchArea;
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[22, 4 + i] = analogs[i]?.Area ?? 0;

                // Строка 23: Корректировка на площадь (ФОРМУЛА)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[23, 4 + i].Formula =
                        $"=IF({col}22>0, MAX(-0.15, MIN(0.15, (($C$22/{col}22)^-0.11)-1)), 0)";
                }

                // Строка 24: Скорректированная стоимость (после площади)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[24, 4 + i].Formula = $"={col}21*(1+{col}23)";
                }

                // Строка 25: Расположение на этаже
                worksheet.Cells[25, 3] = $"{researchFloor} из {researchTotalFloors}";
                for (int i = 0; i < 4; i++)
                {
                    if (analogs[i] != null)
                        worksheet.Cells[25, 4 + i] = $"{analogs[i].Floor} из {analogs[i].Total_Floors}";
                    else
                        worksheet.Cells[25, 4 + i] = "нет данных";
                }

                // Строка 26: Корректировка на этаж
                double objCoeff = researchFloor == 1 ? 0.94 : (researchFloor == researchTotalFloors ? 0.92 : 1);
                for (int i = 0; i < 4; i++)
                {
                    if (analogs[i] != null)
                    {
                        double analogCoeff = analogs[i].Floor == 1 ? 0.94 :
                                            (analogs[i].Floor == analogs[i].Total_Floors ? 0.92 : 1);
                        double correction = (objCoeff / analogCoeff) - 1;
                        worksheet.Cells[26, 4 + i] = correction;
                    }
                    else
                    {
                        worksheet.Cells[26, 4 + i] = 0;
                    }
                }

                // Строка 27: Скорректированная стоимость (после этажа)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[27, 4 + i].Formula = $"={col}24*(1+{col}26)";
                }

                // Строка 28: Состояние отделки
                worksheet.Cells[28, 3] = GetFinishing(realty.Finishing);
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[28, 4 + i] = GetFinishing(analogs[i]?.Finishing);

                // Строка 29: Корректировка на отделку (ФОРМУЛА)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[29, 4 + i].Formula =
                        $"=IF($C$28=\"Без отделки\", " +
                        $"IF(${col}$28=\"Черновая\", -0.05, " +
                        $"IF(${col}$28=\"Чистовая\", -0.07, 0)), " +
                        $"IF($C$28=\"Черновая\", " +
                        $"IF(${col}$28=\"Без отделки\", 0.05, " +
                        $"IF(${col}$28=\"Чистовая\", -0.02, 0)), " +
                        $"IF($C$28=\"Чистовая\", " +
                        $"IF(${col}$28=\"Без отделки\", 0.07, " +
                        $"IF(${col}$28=\"Черновая\", 0.02, 0)), 0)))";
                }

                // Строка 30: Скорректированная стоимость (после отделки)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[30, 4 + i].Formula = $"={col}27*(1+{col}29)";
                }

                // Строка 31: Наличие мебели
                worksheet.Cells[31, 3] = GetFurniture(realty.Furniture);
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[31, 4 + i] = GetFurniture(analogs[i]?.Furniture);

                // Строка 32: Корректировка на мебель (ФОРМУЛА)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[32, 4 + i].Formula =
                        $"=IF(AND(${col}$31=\"Да\", $C$31=\"Нет\"), -0.02, " +
                        $"IF(AND(${col}$31=\"Нет\", $C$31=\"Да\"), 0.02, 0))";
                }

                // Строка 33: Скорректированная стоимость (после мебели)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[33, 4 + i].Formula = $"={col}30*(1+{col}32)";
                }

                // Строка 34: Наличие балкона
                worksheet.Cells[34, 3] = GetBalcony(realty.Balcony);
                for (int i = 0; i < 4; i++)
                    worksheet.Cells[34, 4 + i] = GetBalcony(analogs[i]?.Balcony);

                // Строка 35: Корректировка на балкон (ФОРМУЛА)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[35, 4 + i].Formula =
                        $"=IF(AND(${col}$34=\"имеется\", $C$34=\"не имеется\"), -0.02, " +
                        $"IF(AND(${col}$34=\"не имеется\", $C$34=\"имеется\"), 0.02, 0))";
                }

                // Строка 36: Скорректированная стоимость (после балкона)
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[36, 4 + i].Formula = $"={col}33*(1+{col}35)";
                }

                // Строка 37: Итоговая скорректированная стоимость
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[37, 4 + i].Formula = $"={col}36";
                }

                // Строка 39: Коэффициент вариации
                worksheet.Cells[39, 3].Formula = "=STDEV(D37:G37)/AVERAGE(D37:G37)";

                // Строка 40: Общая валовая коррекция
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[40, 4 + i].Formula =
                        $"=ABS({col}14)+ABS({col}20)+ABS({col}23)+ABS({col}26)+ABS({col}29)+ABS({col}32)+ABS({col}35)";
                }

                // Строка 41: Величина обратная валовой коррекции
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[41, 4 + i].Formula = $"=1/{col}40";
                }

                // Строка 42: Сумма показателей
                worksheet.Cells[42, 3].Formula = "=SUM(D41:G41)";

                // Строка 43: Удельный вес аналога
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[43, 4 + i].Formula = $"={col}41/$C$42";
                }

                // Строка 44: Расчетная цена аналогов
                for (int i = 0; i < 4; i++)
                {
                    char col = (char)((int)'D' + i);
                    worksheet.Cells[44, 4 + i].Formula = $"={col}37*{col}43";
                }

                // Строка 45: Рыночная стоимость за 1 кв.м
                worksheet.Cells[45, 3].Formula = "=SUM(D44:G44)";

                // Строка 46: Стоимость объекта оценки
                worksheet.Cells[46, 3].Formula = "=$C$45*$C$5";

                // Строка 47: Стоимость с округлением
                worksheet.Cells[47, 3].Formula = "=ROUND($C$46, 0)";

                workbook.Save();
                workbook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка при заполнении Excel шаблона: {ex.Message}");
            }
            finally
            {
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ExportToWord(Expertise selectedExpertise)
        {
            string projectRoot = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string templatePath = System.IO.Path.Combine(projectRoot, "Export", "Шаблон.docx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Шаблон не найден по пути: {templatePath}");
            }

            string fileName = GenerateFileName(selectedExpertise.Realties) + "_Заключение.docx";
            string exportFolder = System.IO.Path.Combine(projectRoot, "Export");

            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
            }

            string outputPath = System.IO.Path.Combine(exportFolder, fileName);

            int fileNumber = 1;
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(fileName);
            string extension = System.IO.Path.GetExtension(fileName);

            while (File.Exists(outputPath))
            {
                string newFileName = $"{fileNameWithoutExtension}_{fileNumber}{extension}";
                outputPath = System.IO.Path.Combine(exportFolder, newFileName);
                fileNumber++;
            }

            File.Copy(templatePath, outputPath, true);
            FillWordDocument(outputPath, selectedExpertise);
        }

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

                var replacements = GetWordReplacements(expertise);
                foreach (var replacement in replacements)
                {
                    ReplaceTextInWordDocument(doc, replacement.Key, replacement.Value);
                }

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

        private Dictionary<string, string> GetWordReplacements(Expertise expertise)
        {
            var realty = expertise.Realties;
            var caseInfo = expertise.Case;
            var expert = expertise.Experts;

            string dateStart = expertise.Date_Start?.ToString("dd.MM.yyyy") ?? "_____";
            string dateEnd = expertise.Date_End?.ToString("dd.MM.yyyy") ?? "_____";
            string dateCase = caseInfo?.Date_Case.ToString("dd.MM.yyyy") ?? "_____";
            string fullAddress = realty?.FullAddress ?? "Адрес не указан";

            string purposeText = "";
            if (realty?.Purpose != null)
            {
                int rooms = realty.Purpose;
                purposeText = $"{rooms}-комнатная";
            }

            string wallMaterial = GetWallMaterialForWord(realty?.Wall_Material);
            string finishing = GetFinishingForWord(realty?.Finishing);
            string expertFIO = expert?.FIO ?? "_________________________";
            string caseNumber = caseInfo?.Number_Case ?? "_____";
            string court = caseInfo?.court ?? "______________";
            string price = realty?.Price.HasValue == true ? $"{Math.Round(realty.Price.Value, 0):N0}" : "_________";
            string Baclony = realty.Balcony;
            string Furniture = realty.Furniture == "Да" ? "Имеется" : "Не имеется";
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
                { "[Balcony]", Baclony },
                { "[Furniture]", Furniture },
                { "[Wall_Material]", wallMaterial },
                { "[Price]", price }
            };
        }

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

        private string GenerateFileName(Realties realty)
        {
            if (realty == null)
                return $"Экспертиза_{DateTime.Now:yyyyMMdd_HHmmss}";

            string address = "";
            if (!string.IsNullOrWhiteSpace(realty.NameStreet))
                address += realty.NameStreet;
            if (!string.IsNullOrWhiteSpace(realty.House))
                address += "_" + realty.House;
            if (!string.IsNullOrWhiteSpace(realty.Apartment_number))
                address += "_кв" + realty.Apartment_number;

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                address = address.Replace(c, '_');
            }
            address = address.Replace(" ", "_");

            return string.IsNullOrWhiteSpace(address) ? $"Экспертиза_{DateTime.Now:yyyyMMdd_HHmmss}" : $"Экспертиза_{address}";
        }

        private void AddCase_Click(object sender, RoutedEventArgs e)
        {
            var caseWindow = new AddEditCase(null);
            caseWindow.Owner = Window.GetWindow(this);

            if (caseWindow.ShowDialog() == true)
            {
                RefreshCasesList();
                if (caseWindow.DataContext is Case newCase && newCase != null)
                {
                    CaseComboBox.SelectedValue = newCase.Number_Case;
                    _currentExpertise.ID_Case = newCase.Number_Case;
                }
            }
        }

        private void EditCase_Click(object sender, RoutedEventArgs e)
        {
            if (CaseComboBox.SelectedItem == null)
            {
                MessageBox.Show("Выберите дело для редактирования", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedCase = CaseComboBox.SelectedItem as Case;
            if (selectedCase == null)
            {
                MessageBox.Show("Не удалось получить данные выбранного дела", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var caseWindow = new AddEditCase(selectedCase);
            caseWindow.Owner = Window.GetWindow(this);

            if (caseWindow.ShowDialog() == true)
            {
                RefreshCasesList();
                CaseComboBox.SelectedValue = selectedCase.Number_Case;
            }
        }

        private void RefreshCasesList()
        {
            var context = ExpertiseEntities.GetContext();
            var cases = context.Case.ToList()
                .OrderBy(c => c.Number_Case)
                .ToList();

            var currentSelection = CaseComboBox.SelectedValue;

            CaseComboBox.ItemsSource = cases;
            CaseComboBox.DisplayMemberPath = "Number_Case";
            CaseComboBox.SelectedValuePath = "Number_Case";

            if (currentSelection != null)
            {
                CaseComboBox.SelectedValue = currentSelection;
            }
        }

        // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ДЛЯ ЭКСПОРТА ====================

        private string GetWallMaterialForExcel(string material)
        {
            if (string.IsNullOrEmpty(material)) return "нет данных";
            return material.ToLower() == "кирпич" ? "кирпич" : material;
        }

        private string GetFullAddress(Realties realty) => realty?.FullAddress ?? "нет данных";
        private string GetFullAddress(Analog_Properties analog) => analog?.FullAddress_Analog ?? "нет данных";
        private string FormatDate(DateTime? date) => date.HasValue ? date.Value.ToString("dd.MM.yyyy") : "нет данных";
        private string GetFinishing(string finishing) => string.IsNullOrEmpty(finishing) ? "нет данных" : finishing;
        private string GetFurniture(string furniture) => string.IsNullOrEmpty(furniture) ? "нет данных" : furniture;
        private string GetBalcony(string balcony) => string.IsNullOrEmpty(balcony) ? "нет данных" : balcony;

        private string GetWallMaterialForWord(string material)
        {
            if (string.IsNullOrEmpty(material)) return "Данные отсутствуют";
            switch (material.ToLower())
            {
                case "кирпич": return "кирпич";
                case "панель": return "панель";
                case "дерево": return "дерево";
                default: return material;
            }
        }

        private string GetFinishingForWord(string finishing)
        {
            if (string.IsNullOrEmpty(finishing)) return "Данные отсутствуют";
            switch (finishing.ToLower())
            {
                case "черновая": return "черновая отделка";
                case "чистовая": return "чистовая отделка";
                default: return finishing;
            }
        }

        // ==================== МЕТОДЫ КОРРЕКТИРОВОК ====================

        private double GetWallMaterialCorrection(string researchMaterial, string analogMaterial)
        {
            if (string.IsNullOrEmpty(researchMaterial) || string.IsNullOrEmpty(analogMaterial)) return 0;
            if (researchMaterial.Equals(analogMaterial, StringComparison.OrdinalIgnoreCase)) return 0;

            var corrections = new Dictionary<string, Dictionary<string, double>>
            {
                { "Кирпич", new Dictionary<string, double> { { "Панель", -0.05 }, { "Дерево", -0.08 } } },
                { "Панель", new Dictionary<string, double> { { "Кирпич", 0.05 }, { "Дерево", -0.03 } } },
                { "Дерево", new Dictionary<string, double> { { "Кирпич", 0.08 }, { "Панель", 0.05 } } }
            };
            if (corrections.ContainsKey(researchMaterial) && corrections[researchMaterial].ContainsKey(analogMaterial))
                return corrections[researchMaterial][analogMaterial];
            return 0.03;
        }

        private double CalculateAreaCorrection(double researchArea, double analogArea)
        {
            if (analogArea <= 0) return 0;
            double ratio = researchArea / analogArea;
            double correction = Math.Pow(ratio, -0.11) - 1;
            return Math.Max(-0.15, Math.Min(0.15, correction));
        }

        private double CalculateFloorCorrection(int researchFloor, int analogFloor, int researchTotalFloors, int analogTotalFloors)
        {
            double researchCoeff = GetFloorCoefficient(researchFloor, researchTotalFloors);
            double analogCoeff = GetFloorCoefficient(analogFloor, analogTotalFloors);
            return (researchCoeff / analogCoeff) - 1;
        }

        private double GetFloorCoefficient(int floor, int totalFloors)
        {
            if (floor == 1) return 0.94;
            if (floor == totalFloors) return 0.92;
            return 1;
        }

        private double GetFinishingCorrection(string researchFinishing, string analogFinishing)
        {
            if (string.IsNullOrEmpty(researchFinishing) || string.IsNullOrEmpty(analogFinishing)) return 0;
            var corrections = new Dictionary<string, Dictionary<string, double>>
            {
                { "Без отделки", new Dictionary<string, double> { { "Без отделки", 0.0 }, { "Черновая", -0.05 }, { "Чистовая", -0.07 } } },
                { "Черновая", new Dictionary<string, double> { { "Без отделки", 0.05 }, { "Черновая", 0.0 }, { "Чистовая", -0.02 } } },
                { "Чистовая", new Dictionary<string, double> { { "Без отделки", 0.07 }, { "Черновая", 0.02 }, { "Чистовая", 0.0 } } }
            };
            if (corrections.ContainsKey(researchFinishing) && corrections[researchFinishing].ContainsKey(analogFinishing))
                return corrections[researchFinishing][analogFinishing];
            return 0;
        }

        private double GetFurnitureCorrection(string researchFurniture, string analogFurniture)
        {
            if (analogFurniture == "Да" && researchFurniture == "Нет") return -0.02;
            if (analogFurniture == "Нет" && researchFurniture == "Да") return 0.02;
            return 0;
        }

        private double GetBalconyCorrection(string researchBalcony, string analogBalcony)
        {
            if (string.IsNullOrEmpty(researchBalcony) || string.IsNullOrEmpty(analogBalcony)) return 0;
            if (analogBalcony == "имеется" && researchBalcony == "не имеется") return -0.02;
            if (analogBalcony == "не имеется" && researchBalcony == "имеется") return 0.02;
            return 0;
        }
    }
}