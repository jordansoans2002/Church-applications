using System;

namespace Create_PPT_UI.Model
{
    internal class Song
    {
        public Song()
        {

        }
        public Song(string name, string lang1, string txt1)
        {
            songName = name;
            text1 = txt1;
        }
        public Song(string name, string txt1, string txt2, string o)
        {
            songName = name;
            text1 = txt1;
            text2 = txt2;
            orientation = o;
        }
        public string songName { get; set; }
        public string lang1 { get; set; }
        public string lang2 { get; set; }
        public string languages
        {
            get
            {
                if (lang2 != null)
                {
                    return "Languages: " + lang1 + ", " + lang2;
                }
                else
                {
                    return "Language: " + lang1;
                }
            }
        }
        public string text1 { get; set; }
        public string text2 { get; set; } = null;

        public string orientation {  get; set; }
    }
}
