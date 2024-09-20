namespace Create_PPT_UI.Model
{
    internal class Setting
    {
        public Setting(string settingName,string settingValue) {
            this.settingName = settingName;
            this.settingValue = settingValue;
        }
        public string settingName {  get; set; }
        public string settingValue { get; set; }
    }
}
