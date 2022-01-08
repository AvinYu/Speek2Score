﻿using System;
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
        private string excelFileName;
        private Excel.Application excelApp;
        private Excel.Workbook excelWB;
        private Excel.Worksheet excelWS;
        private int cellX, cellY;

        private SpeechRecognizedEventArgs speechEvent;
        private SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("zh-TW"));
        private GrammarBuilder gBuilder = new GrammarBuilder();
        private Choices vocabulary = new Choices();
        private Grammar grammer;

        private int studentNum = 30;
        private string temp;
        private List<string> scoreTable = new List<string> { };

        private DispatcherTimer timer = new DispatcherTimer();

        public MainWindow()
        {
            CheckMicrophone();
            InitializeComponent();
            UISetup();
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
                //Close();
            }
        }

        private void UISetup()
        {
            tbExcelFilename.Text = tbExcelFilenameHint;
            tbExcelFilename.Foreground = Brushes.LightGray;
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

        private void BuildExcelFile()
        {
            currentDirectory = Directory.GetCurrentDirectory();
            excelApp = new Excel.Application();
            excelWB = excelApp.Workbooks.Add();
            excelWS = new Excel.Worksheet();
            excelWS = excelWB.Worksheets[1];
            excelWS.Name = "Sheet1";
            excelWS.Cells[1, 1] = "座號";
            excelWS.Cells[1, 2] = "期末考";

            for (int i = 2; i <= studentNum + 1; i++)
            {
                excelWS.Cells[i, 1] = i - 1;
            }
        }

        private void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            speechEvent = e;
            if (speechEvent.Result.Text.Contains("取消")) // 前次辨識失敗
            {
                tb.Text = "取消成功";
                timer.Start();
            }
            else
            {
                if (temp != null)
                {
                    scoreTable.Add(temp); // debug用，將前次辨識的值存入 顯示大表
                    // X號Y分
                    //split(號)
                    // X   Y分
                    string num, score;
                    try
                    {
                        num = temp.Split(char.Parse("號"))[0];
                        score = (temp.Split(char.Parse("號"))[1]).Split(char.Parse("分"))[0];
                        cellX = Convert.ToInt32(num);
                        //cellY = Convert.ToInt32(score);
                        excelWS.Cells[cellX + 1, 2] = score;
                    }
                    catch { }
                }

                tbScoreTable.Text = String.Join("\n", scoreTable); // debug顯示用

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

        private void SaveExcelFileIfChanged()
        {
            if ((tbExcelFilename.Text != "") && (tbExcelFilename.Text != tbExcelFilenameHint))
            {
                excelFileName = tbExcelFilename.Text;
                excelWB.SaveAs(currentDirectory + "\\" + excelFileName);
            }
        }

        #region UI element events

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
            SaveExcelFileIfChanged();
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            BuildExcelFile();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveExcelFileIfChanged();
            excelWB.Close(SaveChanges: false);
            excelApp.Quit();
        }

        #endregion UI element events
    }
}