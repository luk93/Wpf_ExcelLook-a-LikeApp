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

namespace Wpf_ExcelLook_a_LikeApp
{
    /// <summary>
    /// Interaction logic for PopUp.xaml
    /// </summary>
    public partial class PopUp : Window
    {
        public PopUp()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.row_PopUp = int.Parse(rowTB.Text);
            MainWindow.col_PopUp = int.Parse(colTB.Text);
            MainWindow.wsName_PopUp = wsNameTB.Text;
            this.DialogResult = true;
            this.Close();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void rowTB_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void wsNameTB_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
