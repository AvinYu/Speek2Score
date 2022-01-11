using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
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
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Speak2Score
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string tbExcelFilenameHint = "Excel Filename";
        private string currentDirectory;
        private string excelFullFilePath;
        private Excel.Application excelApp;
        private Excel.Workbook excelWB;
        private Excel.Worksheet excelWS;
        private int cellX, cellY;
        private string num, score;

        private SpeechRecognizedEventArgs speechEvent;
        private SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("zh-TW"));
        private GrammarBuilder gBuilder = new GrammarBuilder();
        private Choices vocabulary = new Choices();
        private Grammar grammer;

        private int studentNum;
        private string temp;

        private List<string> scoreTable = new List<string> { };

        private DispatcherTimer timer = new DispatcherTimer();

        public MainWindow()
        {
            CheckMicrophone();
            InitializeComponent();
            UISetup();
            BindingSetup();
            studentNum = buttonArray.btnNumber;

            BuildVocabulary();
            BuildSpeechRecognition();
            TimerSetup();
        }

        private void BindingSetup()
        {
            Binding btnNum = new Binding()
            {
                Source = buttonArray,
                Path = new PropertyPath("btnNumber"),
                Mode = BindingMode.TwoWay
            };
            tbStudentNumber.SetBinding(TextBox.TextProperty, btnNum);
        }

        private void CheckMicrophone()
        {
            try
            {
                recEngine.SetInputToDefaultAudioDevice();
            }
            catch
            {
                MessageBox.Show("No Microphone Detected, Please Check");
                Close();
            }
        }

        private void UISetup()
        {
            tbExcelFilename.Text = tbExcelFilenameHint;
            tbExcelFilename.Foreground = Brushes.LightGray;
        }

        private void BuildVocabulary()
        {
            for (int i = 1; i <= studentNum; i++)
            {
                for (int j = 0; j <= 100; j++)
                {
                    vocabulary.Add(String.Concat(i.ToString(), "號", j.ToString()));
                }
            }

            vocabulary.Add("取 消");
            vocabulary.Add("O K");
        }

        private void BuildSpeechRecognition()
        {
            gBuilder.Append(vocabulary, 0, 1);
            grammer = new Grammar(gBuilder);
            recEngine.LoadGrammarAsync(grammer);
            recEngine.SpeechRecognized += recEngine_SpeechRecognized;
        }

        private void TimerSetup()
        {
            timer.Tick += new EventHandler(timer_tick);
            timer.Interval = new TimeSpan(0, 0, 2);
        }

        private void BuildExcelFile()
        {
            excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            excelWB = excelApp.Workbooks.Add();
            excelWS = new Excel.Worksheet();
            excelWS = excelWB.Worksheets[1];
            excelWS.Name = "Sheet1";
        }

        private void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            speechEvent = e;

            if (speechEvent.Result.Text.Contains("取消")) // 前次辨識失敗
            {
                tb.Text = "取消成功";
                lbStudent.Foreground = Brushes.Red;
                lbScore.Foreground = Brushes.Red;
                timer.Start();
            }
            else
            {
                try
                {
                    lbStudent.Foreground = Brushes.Black;
                    lbScore.Foreground = Brushes.Black;

                    lbStudent.Content = speechEvent.Result.Text.Split(char.Parse("號"))[0];
                    lbScore.Content = speechEvent.Result.Text.Split(char.Parse("號"))[1];
                }
                catch { }

                if (temp != null)
                {
                    scoreTable.Add(temp); // debug用，將前次辨識的值存入 顯示大表

                    try
                    {
                        num = temp.Split(char.Parse("號"))[0];
                        score = temp.Split(char.Parse("號"))[1];
                        cellX = Convert.ToInt32(num);
                        excelWS.Cells[cellX + 1, 2] = score;
                        //TODO: Same file if cellY not empty, go to cellY+1
                    }
                    catch { }
                }

                tbScoreTable.Text = String.Join("\n", scoreTable); // debug顯示用
                //tb.Text = speechEvent.Result.Text; // 改成大字幕

                temp = speechEvent.Result.Text;

                if (speechEvent.Result.Text == "OK")
                {
                    tb.Text = "OK";
                    lbStudent.Content = "";
                    lbScore.Content = "";
                    RecognizeAsync(false);
                    SaveExcelFile();
                }
            }
        }

        private void timer_tick(object sender, EventArgs e)
        {
            timer.Stop();
            if (speechEvent.Result.Text.Contains("取消"))
            {
                tb.Text = "";
                temp = null;
            }
        }

        private void SaveExcelFile()
        {
            if ((tbExcelFilename.Text != "") && (tbExcelFilename.Text != tbExcelFilenameHint))
            {
                excelWS.Cells[1, 1] = "座號";
                excelWS.Cells[1, 2] = "期末考";
                for (int i = 2; i <= studentNum + 1; i++)
                {
                    excelWS.Cells[i, 1] = i - 1;
                }

                currentDirectory = Directory.GetCurrentDirectory();
                excelFullFilePath = String.Concat(currentDirectory, "\\", tbExcelFilename.Text);
                excelWB.SaveAs(excelFullFilePath);
                tb.Text = "File Saved Success";
            }
        }

        private void RecognizeAsync(bool isTrue)
        {
            if (isTrue)
            {
                try
                {
                    recEngine.RecognizeAsync(RecognizeMode.Multiple);
                }
                catch { };
            }
            else
            {
                try
                {
                    recEngine.RecognizeAsyncStop();
                }
                catch { };
            }

            btnStart.IsEnabled = !isTrue;
            btnStop.IsEnabled = isTrue;
        }

        #region UI element events

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            RecognizeAsync(true);
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            RecognizeAsync(false);
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            tb.Text = "";
            tbScoreTable.Text = "";
            scoreTable.Clear();
            temp = null;
        }

        private void BtnRecover_Click(object sender, RoutedEventArgs e)
        {
            tb.Text = "Recover Success";
            temp = null;
        }

        private void tbExcelFilename_LostFocus(object sender, RoutedEventArgs e)
        {
            if (tbExcelFilename.Text == "")
            {
                tbExcelFilename.Text = tbExcelFilenameHint;
                tbExcelFilename.Foreground = Brushes.LightGray;
            }
        }

        private void tbExcelFilename_GotFocus(object sender, RoutedEventArgs e)
        {
            if (tbExcelFilename.Text == tbExcelFilenameHint)
            {
                tbExcelFilename.Text = "";
                tbExcelFilename.Foreground = Brushes.Black;
            }
        }

        private void BtnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            SaveExcelFile();
        }

        private void BtnStudentplus_Click(object sender, RoutedEventArgs e)
        {
            if (studentNum < 50)
            {
                studentNum++;
                buttonArray.ButtonNumberChange(studentNum);
            }
        }

        private void BtnStudentMinus_Click(object sender, RoutedEventArgs e)
        {
            if (studentNum > 1)
            {
                studentNum--;
                buttonArray.ButtonNumberChange(studentNum);
            }
        }

        private void tbStudentNumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            while (Convert.ToInt32(tbStudentNumber.Text) != studentNum)
            {
                if (Convert.ToInt32(tbStudentNumber.Text) > studentNum) studentNum++;
                else studentNum--;

                buttonArray.ButtonNumberChange(studentNum);
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            BuildExcelFile();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveExcelFile();
            excelWB.Close(SaveChanges: false);
            excelApp.Quit();

            Marshal.ReleaseComObject(excelWS);
            Marshal.ReleaseComObject(excelWB);
            Marshal.ReleaseComObject(excelApp);
        }

        #endregion UI element events
    }
}