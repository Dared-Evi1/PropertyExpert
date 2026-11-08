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
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Net.Mime.MediaTypeNames;

namespace Diplom
{

        /// <summary>
    /// Логика взаимодействия для Registrtion.xaml
    /// </summary>
    public partial class Registrtion : Page
    {
        public Experts expert;
        public List<Experts> currentExperts;
        public Registrtion()
        {
            InitializeComponent();
        }

        private void Welcome_Click(object sender, RoutedEventArgs e)
        {
            string login = Login.Text;
            string password = Password.Password;
            if (login == "" || password == "")
            {
                MessageBox.Show("Есть пустые поля");
                return;
            }

            expert = ExpertiseEntities.GetContext().Experts.ToList().Find(p => p.Login == login && p.Password == password);
            if (login == "admin" && password == "admin")
            {

                Manager.MainFrame.Navigate(new ExpertPage(expert));
            }
            else
            {
                if (expert != null)
                {
                    if (expert.Status == "Работает")
                    {
                        Manager.MainFrame.Navigate(new ExpertisePage(expert));
                    }
                    else
                    {
                        MessageBox.Show("Вы не можете зайти, так как уволены");

                    }
                }
                else
                {

                    MessageBox.Show("Введены неверные данные");
                }
            }
        }

        


    }
}
