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
    /// <summary>
    /// Логика взаимодействия для YearSelectionDialog.xaml
    /// </summary>
    public partial class YearSelectionDialog : Window
    {
        public int SelectedYear { get; private set; }
        public YearSelectionDialog()
        {
            InitializeComponent();

            // Заполняем года (последние 10 лет)
            int currentYear = DateTime.Now.Year;
            for (int year = currentYear; year >= currentYear - 10; year--)
            {
                YearComboBox.Items.Add(year);
            }

            // Выбираем текущий год по умолчанию
            YearComboBox.SelectedItem = currentYear;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (YearComboBox.SelectedItem != null)
            {
                SelectedYear = (int)YearComboBox.SelectedItem;
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите год", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
