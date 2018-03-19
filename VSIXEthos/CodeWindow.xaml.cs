namespace VSIXEthos
{
    /// <summary>
    /// Interaction logic for CodeWindow.xaml
    /// </summary>
    public partial class CodeWindow
    {
        public CodeWindow()
        {
            InitializeComponent();
        }

        public CodeWindow(CodeWindowViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
        }
    }
}
