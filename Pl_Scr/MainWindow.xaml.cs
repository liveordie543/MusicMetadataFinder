using System;
using System.Windows;
using Microsoft.Win32;

namespace Pl_Scr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SearchBtn_Click(object sender, RoutedEventArgs e)
        {
            OutputTxtBox.Text = LastFmApiHelper.GetReleaseDateBySong(SongTitleTxtBox.Text, ArtistTxtBox.Text);
        }

        private void BrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };
            bool? result = dialog.ShowDialog();
            if (result == true)
            {
                PlaylistFileTxtBox.Text = dialog.FileName;
            }
        }

        private void ProcessFileBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                new XlsxProcessor(PlaylistFileTxtBox.Text).Process();
                OutputTxtBox.Text = "Playlist successfully processed.";
            }
            catch (Exception ex)
            {
                OutputTxtBox.Text = ex.Message + Environment.NewLine + ex.StackTrace;
            }
        }
    }
}
