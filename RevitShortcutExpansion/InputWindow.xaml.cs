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

namespace RevitShortcutExpansion
{
    /// <summary>
    /// InputWindow.xaml 的交互逻辑
    /// </summary>
    public partial class InputWindow : Window
    {
        private string _suggestedName;

        public string SaveName { get; private set; }
        public InputWindow(string suggestedName)
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            _suggestedName = suggestedName;
            inputName.Text = _suggestedName;
            inputName.SelectionStart = inputName.Text.Length;
            inputName.Focus();

            Ok.Click += Ok_Click;
            Cancel.Click += Cancel_Click;
        }
        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            SaveName = inputName.Text;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
