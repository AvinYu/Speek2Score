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

namespace Speak2Score
{
    /// <summary>
    /// ButtonArray.xaml 的互動邏輯
    /// </summary>
    public partial class ButtonArray : UserControl
    {
        private const int btnWidth = 24;
        private const int btnHeight = 18;

        public static readonly DependencyProperty btnNumberValueProperty = DependencyProperty.Register(nameof(btnNumber), typeof(int), typeof(ButtonArray), new PropertyMetadata());

        public int btnNumber
        {
            get => (int)GetValue(btnNumberValueProperty);
            set => SetValue(btnNumberValueProperty, value);
        }

        private static readonly DependencyProperty dynamicWidthValueProperty = DependencyProperty.Register(nameof(dynamicWidth), typeof(int), typeof(ButtonArray), new PropertyMetadata());

        private int dynamicWidth
        {
            get => (int)GetValue(dynamicWidthValueProperty);
            set => SetValue(dynamicWidthValueProperty, value);
        }

        private double lastRow;

        public ButtonArray()
        {
            btnNumber = 30;
            WidthChange(btnNumber);

            InitializeComponent();
            BuildButtons();
        }

        private void BuildButtons()
        {
            for (int i = 1; i <= btnNumber; i++)
            {
                Button btn = new Button();
                btn.Name = "btn" + i.ToString();
                wrappanel.Children.Add(btn);
            }
        }

        public void ButtonNumberChange(int studentnumber)
        {
            if (studentnumber > btnNumber)
            {
                Button btn = new Button();
                btn.Name = "btn" + studentnumber.ToString();
                wrappanel.Children.Add(btn);
                btnNumber++;
            }
            else
            {
                wrappanel.Children.RemoveAt(studentnumber);
                btnNumber--;
            }
            WidthChange(btnNumber);
        }

        private void WidthChange(int btnnumber)
        {
            lastRow = Math.Ceiling((double)btnnumber / 5);
            dynamicWidth = btnWidth * (int)lastRow;
        }
    }
}