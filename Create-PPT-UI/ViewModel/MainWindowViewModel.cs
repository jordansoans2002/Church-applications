using Create_PPT_UI.Model;
using Create_PPT_UI.MVVM;
using System.Collections.ObjectModel;

namespace Create_PPT_UI.ViewModel
{
    internal class MainWindowViewModel : ViewModelBase
    {
        public ObservableCollection<Song> Songs {  get; set; }
        public ObservableCollection<Setting> Settings { get; set; }
        public RelayCommand AddSongCommand => new RelayCommand(execute => AddSong(), canExecute => { return true; });

        public MainWindowViewModel()
        {
            Songs = new ObservableCollection<Song>();
            Settings = new ObservableCollection<Setting>();

            Settings.Add(new Setting("Lyrics path", "C:/Users/admin/Desktop/Church"));

            AddSong();
            AddSong();
            AddSong();
           
        }

        
        private Song selectedSong;
        public Song SelectedSong
        {
            get { return selectedSong; }
            set 
            {
                selectedSong = value;
                OnPropertyChanged();
            }
        }

        private void AddSong()
        {
            Songs.Add(new Song
            {
                songName = "The Steadfast Love of the Lord",
                lang1 = "english",
                text1 = "The steadfast love of the Lord never ceases, \n" +
                "His mercies never come to an end \n" +
                "They are new every morning, new every morning \n" +
                "Great is Thy faithfulness O Lord \n" +
                "Great is Thy faithfulness \n"
            });
           // Songs.Add(new Song(
           //    "Asha Meri",
           //    "english",
           //    "Humme wo doori,\n" +
           //    "Thi kitni gehri\n" +
           //    "Kitna bada tha, wo fasla\n" +
           //    "Mayus hokar, swarg ki or dekha\n" +
           //    "Nirasha me tera naam liya\n"
           //));
        }
    }
}
