using Create_PPT_UI.ViewModel;
using System.Windows;
using System.Windows.Controls;

namespace Create_PPT_UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainWindowViewModel viewModel = new MainWindowViewModel();
            DataContext = viewModel;
        }


        private void ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if(sender == lang1_textbox)
            {
                lang2_textbox.ScrollToVerticalOffset(e.VerticalOffset);
            }
            else
            {
                lang1_textbox.ScrollToVerticalOffset(e.VerticalOffset);
            }
        }
    }

}
