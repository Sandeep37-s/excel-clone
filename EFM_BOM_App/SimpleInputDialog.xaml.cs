using System.Windows;

namespace EFM_BOM_App
{
    public partial class SimpleInputDialog : Window
    {
        public string InputValue => TxtInput.Text;

        public SimpleInputDialog(string title, string prompt, string defaultValue = "")
        {
            InitializeComponent();
            Title = title;
            LblPrompt.Text = prompt;
            TxtInput.Text = defaultValue;
            TxtInput.SelectAll();
            TxtInput.Focus();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}

