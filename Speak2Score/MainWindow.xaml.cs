using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

namespace Speak2Score
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private SpeechRecognizedEventArgs speechEvent;
        private SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("zh-TW"));
        private GrammarBuilder gBuilder = new GrammarBuilder();
        private Choices vocabulary = new Choices();
        private Grammar grammer;
        private int studentNum = 50;
        private string temp;
        private List<string> ScoreTable = new List<string> { };
        private DispatcherTimer timer = new DispatcherTimer();

        public MainWindow()
        {
            CheckMicrophone();
            InitializeComponent();
            BuildVocabulary();
            BuildSpeechRecognition();
            TimerSetup();
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

        private void BuildVocabulary()
        {
            //for (int i = 1; i <= studentNum; i++)
            //{
            //    vocabulary.Add(String.Concat(i.ToString(), "號"));
            //}
            //for (int i = 0; i <= 100; i++)
            //{
            //    vocabulary.Add(String.Concat(i.ToString(), "分"));
            //}

            for (int i = 1; i <= studentNum; i++)
            {
                for (int j = 0; j <= 100; j++)
                {
                    vocabulary.Add(String.Concat(i.ToString(), "號", j.ToString(), "分"));
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

        private void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            speechEvent = e;
            if (speechEvent.Result.Text.Contains("取消")) // 前次辨識失敗
            {
                tb.Text = "取消成功";
                timer.Start();
            }
            else if (speechEvent.Result.Text.Contains("OK")) // 最後一名學生辨識正確後 OK
            {
                tb.Text = "OK";
                ScoreTable.Add(temp); // 將前次辨識的值存入大表
            }
            else
            {
                if (temp != null) ScoreTable.Add(temp);
                tbScoreTable.Text = String.Join("\n", ScoreTable); // debug顯示用

                tb.Text = speechEvent.Result.Text;
                temp = speechEvent.Result.Text;
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

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
            btnStart.IsEnabled = false;
            btnStop.IsEnabled = true;
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            recEngine.RecognizeAsyncStop();
            btnStart.IsEnabled = true;
            btnStop.IsEnabled = false;
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            tb.Text = "";
            tbScoreTable.Text = "";
            ScoreTable.Clear();
            temp = null;
        }

        private void BtnRecover_Click(object sender, RoutedEventArgs e)
        {
            tb.Text = "Recover Success";
            temp = null;
        }
    }
}