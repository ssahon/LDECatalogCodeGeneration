using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.VisualStudio.PlatformUI;
using VSIXEthos.Properties;

namespace VSIXEthos
{
    public class CodeWindowViewModel: INotifyPropertyChanged
    {
        private string _catalogCode;
        private ICommand _copyCommand;

        public CodeWindowViewModel()
        {
            _copyCommand=new DelegateCommand(OnCopy);
        }

        private void OnCopy(object obj)
        {
            Clipboard.Clear();
            Clipboard.SetText(_catalogCode); 
        }

        public string CatalogCode
        {
            get { return _catalogCode; }
            set { _catalogCode = value; OnPropertyChanged();}
        }

        public ICommand CopyCommand
        {
            set { _copyCommand = value; }
            get { return _copyCommand; }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
