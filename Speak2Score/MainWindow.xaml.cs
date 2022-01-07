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

namespace Speak2Score
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        private GrammarBuilder gBuilder = new GrammarBuilder();
        private Choices vocabulary = new Choices();
        private Grammar grammer;
        private int studentNum = 50;
        private string temp;
        private List<string> ScoreTable = new List<string> { };

        public MainWindow()
        {
            CheckMicrophone();
            InitializeComponent();
            BuildVocabulary();
            BuildSpeechRecognition();
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
            for (int i = 1; i <= studentNum; i++)
            {
                vocabulary.Add(String.Concat(i.ToString(), "號"));
            }
            for (int i = 0; i <= 100; i++)
            {
                vocabulary.Add(String.Concat(i.ToString(), "分"));
            }
            vocabulary.Add("復原");
            vocabulary.Add("OK");
        }

        private void BuildSpeechRecognition()
        {
            gBuilder.Append(vocabulary, 0, 2);

            grammer = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammer);

            recEngine.SpeechRecognized += recEngine_SpeechRecognized;
        }

        private void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text == "復原")
            {
                tb.Text = "";
                temp = "";
            }
            else
            {
                ScoreTable.Add(temp);
                tbScoreTable.Text = ScoreTable.ToString();

                tb.Text = e.Result.Text;
                temp = e.Result.Text;
            }
        }

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            recEngine.RecognizeAsyncStop();
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            tb.Text = "";
            tb.Text = "";
            ScoreTable.Clear();
        }

        private void BtnRecover_Click(object sender, RoutedEventArgs e)
        {
            tb.Text = "";
            temp = "";
        }
    }
}