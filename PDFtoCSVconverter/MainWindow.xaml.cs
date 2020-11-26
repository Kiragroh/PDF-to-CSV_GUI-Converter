using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Runtime.InteropServices;
using PDFtoCSVconverter.Properties;
using JR.Utils.GUI.Forms;
using System.Drawing;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using itextPdfTextCoordinates;



namespace PDFtoCSVconverter
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string logtext;
        public string inputFolderPath;
        public string outputFolderPath;
        public string defaultFolderPath;
        //public string outputFilename;
        //public StreamWriter reportFile;
        // holds standard error output collected during run of the DCMTK script
        private static StringBuilder stdErr = new StringBuilder("");
        

        public MainWindow()
        {
            InitializeComponent();
            // Add existing WPF control to the script window.
            //var mainControl = new EclipseDataMiner.MainWindow();
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string scriptVersion = fvi.FileVersion;
            //Window window = Window.GetWindow(this);
            Title = "PDF to CSV-Converter by MG (v." + scriptVersion + ")";


            ProgressTextBlock.Text = "Press the Convert-Button to begin PDF-To-CSV";
            progressBar.Value = 0;

            
            if (Directory.Exists(Settings.Default.Input))
            {
                inputFolderPath = Settings.Default.Input;
            }
            else
            {
                inputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            if (Directory.Exists(Settings.Default.Output))
            {
                outputFolderPath = Settings.Default.Output;
            }
            else
            {
                outputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+@"\csv";
                
            }
            
            pathInputTextBlock.Text = inputFolderPath;
            Settings.Default.Input = inputFolderPath;
            pathOutputTextBlock.Text = outputFolderPath;
            Settings.Default.Output = outputFolderPath;


            ShowLogMsg("Input path: " + inputFolderPath);
            //ShowLogMsg("Output path: " + outputFolderPath);

        }
        private async void runButton_Click(object sender, RoutedEventArgs e)
        {

            if (Directory.Exists(inputFolderPath))
            {

                string dir = inputFolderPath;
                var toRead = Enumerable.Empty<string>();
                if (subDirCheckBox.IsChecked == false)
                {
                    toRead = Directory.GetFiles(dir).Where(f => Path.GetFileName(f).ToLower().Contains(".pdf"));
                }
                else
                {
                    toRead = Directory.GetFiles(dir, "*.pdf", SearchOption.AllDirectories).Where(f => Path.GetFileName(f).Contains(".pdf"));
                }
                int np = toRead.Count();
                ShowLogMsg("\nNew Conversion has started.");
                ShowLogMsg("Number of PDF files for conversion: " + np.ToString());

                if (np > 0)
                {
                    ShowLogMsg("");
                    int count = 1;
                    //List<string> peakList = new List<string>();
                    foreach (var file in toRead)
                    {
                        try
                        {
                            PdfReader pdfReader = new PdfReader(file);
                            int numberOfPages = pdfReader.NumberOfPages;
                            int pageNo = 1;
                            //float[] limitCoordinates = { 52, 671, 357, 728 };//{LowerLeftX,LowerLeftY,UpperRightX,UpperRightY}
                            float[] limitCoordinates = null;

                            List<string> wholePDFstringList_ = new List<string>();

                            for (int i = 0; i < numberOfPages; i++)
                            {
                                // This line gives the lists of rows consisting of one or more columns
                                //if you pass the third parameter as null the it returns the content for whole page
                                // but if you pass the coordinates then it returns the content for that coords only
                                var lineText = LineUsingCoordinates.getLineText(file, pageNo, limitCoordinates);

                                // For detecting the table we are using the fact that the 'lineText' item which length is 
                                // less than two is surely not the part of the table and the item which is having more than
                                // 2 elements is the part of table
                                foreach (var row in lineText)
                                {
                                    string rowString = "";
                                    if (row.Count > 0)
                                    {
                                        for (var col = 0; col < row.Count; col++)
                                        {
                                            string trimmedValue = row[col].Trim();
                                            if (trimmedValue != "")
                                            {
                                                //Console.Write("|" + trimmedValue + "|");
                                                rowString += "|" + trimmedValue + "|";
                                            }
                                        }
                                        //specific code for Elements-Reports
                                        if (rowString.StartsWith("|Object name||Constraint||Objective||Slider positions"))
                                        {
                                            break;
                                        }


                                        if (!rowString.StartsWith("|Created ") & !rowString.StartsWith("|© by Brainlab AG") & !rowString.StartsWith("|["))
                                            wholePDFstringList_.Add(rowString.Replace("||", "zzght99").Replace("|", "").Replace("zzght99", "||"));

                                    }
                                }
                                //Console.ReadLine();
                                var wholePDFstringList = wholePDFstringList_.Distinct();
                                await Task.Run(() => pageNo++);

                                string outputfilePath;

                                if (customSuffix.IsChecked==true)
                                    outputfilePath = outputFolderPath + file.Replace(inputFolderPath, "").ToLower().Replace(".pdf",MakeFilenameValid(CustomSuffixTextBox.Text) +".csv");
                                else
                                {
                                    outputfilePath = outputFolderPath + file.Replace(inputFolderPath, "").ToLower().Replace(".pdf", ".csv");
                                }

                                if (Directory.Exists(Path.GetDirectoryName(outputfilePath)))
                                {
                                    File.WriteAllLines(outputfilePath, wholePDFstringList);
                                }
                                else
                                {
                                    Directory.CreateDirectory(Path.GetDirectoryName(outputfilePath));
                                    File.WriteAllLines(outputfilePath, wholePDFstringList);
                                }

                            }
                            //await Task.Run(() => peakList.Add(DumpPeakDicomTags(file) + "_#"));
                            progressBar.Value = count * 100 / np;
                            ProgressTextBlock.Text = "Processing....";
                            count++;
                        }
                        catch
                        {
                            ShowLogMsg("Error converting " + file);
                            count++;
                        }
                    }
                    //var query = peakList.GroupBy(x => x.ToString(), (y, z) => new { Name = y, Count = z.Count() });

                    // and to test...
                    //foreach (var item in query)
                    //{
                    //    ShowLogMsg(item.Name.ToString()+ item.Count.ToString());
                    //}

                    progressBar.Value = 100;
                    ProgressTextBlock.Text = "PDF-Conversion finished.";
                }
                else
                {
                    ShowLogMsg("No PDF-Conversion possible. Folder does not contain PDF-Files.");
                }
            }
            else
            {
                ShowLogMsg("\nNo PDF-Conversion possible. Folder does not exist.");
            }


        }            

        /// <summary>
        /// ShowLogMsg
        /// </summary>
        /// <param name="dataFile"></param>
        private void ShowLogMsg(string text)
        {
            logTextBox.AppendText(text + "\n");
            logTextBox.SelectionStart = logTextBox.Text.Length;
            logTextBox.ScrollToEnd();
        }

        /// <summary>
        /// window closing event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            //Settings.Default.DaemonTitle = DaemonTitleTextBox.Text;
            Settings.Default.Save();
            base.OnClosing(e);
        }

        /// <summary>
        /// Set input and output filename and folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetInputFolderButton_Click(object sender, RoutedEventArgs e)
        {

            Ookii.Dialogs.Wpf.VistaFolderBrowserDialog fbd = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            fbd.Description = "Select PDF-File input folder";
            //fbd.RootFolder = Environment.SpecialFolder.Favorites;
            fbd.SelectedPath = inputFolderPath;
            fbd.ShowNewFolderButton = true;
            fbd.UseDescriptionForTitle=true;
            
            if (fbd.ShowDialog() == true)
            {
                inputFolderPath = fbd.SelectedPath;
                pathInputTextBlock.Text = inputFolderPath;
                Settings.Default.Input = inputFolderPath;
                ShowLogMsg("\nnew Input path: " + inputFolderPath+"\n");

            }
            
        }

        private void SetOutputFolderButton_Click(object sender, RoutedEventArgs e)
        {

            Ookii.Dialogs.Wpf.VistaFolderBrowserDialog fbd = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            fbd.Description = "Select CSV-File output folder";
            fbd.SelectedPath = outputFolderPath;
            
            fbd.ShowNewFolderButton = true;
            fbd.UseDescriptionForTitle = true;

            if (fbd.ShowDialog() == true)
            {
                outputFolderPath = fbd.SelectedPath;
                pathOutputTextBlock.Text = outputFolderPath;
                Settings.Default.Output = outputFolderPath;
                ShowLogMsg("\nnew Output path: " + outputFolderPath+ "\n");
            }

        }
       

        public string MakeFilenameValid(string s)
        {
            char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (char ch in invalidChars)
            {
                s = s.Replace(ch, '_');
            }
            return s;
        }

       

        private void PeakInputFolderButton_Click(object sender, RoutedEventArgs e)
        {           
            if (Directory.Exists(inputFolderPath) )
            {

                string dir = inputFolderPath;
                var toRead = Enumerable.Empty<string>();
                if (subDirCheckBox.IsChecked == false)
                {
                    toRead = Directory.GetFiles(dir).Where(f => Path.GetFileName(f).ToLower().Contains(".pdf"));
                }
                else
                {
                    toRead = Directory.GetFiles(dir, "*.pdf", SearchOption.AllDirectories).Where(f => Path.GetFileName(f).Contains(".pdf"));
                }
                int np = toRead.Count();
                ShowLogMsg("\nNew Input-Peaking has started.");
                ShowLogMsg("Number of PDF files for peaking: " + np.ToString());

                if (np > 0)
                {
                    ShowLogMsg("");
                    int count = 1;
                    //List<string> peakList = new List<string>();
                    foreach (var file in toRead)
                    {
                        try
                        {
                            //await Task.Run(() => peakList.Add(DumpPeakDicomTags(file) + "_#"));
                            progressBar.Value = count * 100 / np;
                            ProgressTextBlock.Text = "Reading PDF-Files....";
                            count++;
                        }
                        catch
                        {
                            ShowLogMsg("Error reading " + file);
                            count++;
                        }
                    }
                    //var query = peakList.GroupBy(x => x.ToString(), (y, z) => new { Name = y, Count = z.Count() });
                    
                    // and to test...
                    //foreach (var item in query)
                    //{
                    //    ShowLogMsg(item.Name.ToString()+ item.Count.ToString());
                    //}
                    
                    progressBar.Value = 0;
                    ProgressTextBlock.Text = "PDF-Peak finished - Press Start to begin Conversion";
                }
                else
                {
                    //ShowLogMsg("No PDF-Peak possible. Folder does not contain PDF-Files.");
                }
            }
            else
            {
                ShowLogMsg("\nNo PDF-Peak possible. Folder does not exist.");
            }
        }

        private void PeakOutputFolderButton_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(outputFolderPath))
            {

                string dir = outputFolderPath;
                var toRead = Enumerable.Empty<string>();
                if (subDirCheckBox.IsChecked == false)
                {
                    toRead = Directory.GetFiles(dir).Where(f => Path.GetFileName(f).ToLower().Contains(".pdf"));
                }
                else
                {
                    toRead = Directory.GetFiles(dir, "*.pdf", SearchOption.AllDirectories).Where(f => Path.GetFileName(f).Contains(".pdf"));
                }
                int np = toRead.Count();
                ShowLogMsg("\nNew Output-Peaking has started.");
                ShowLogMsg("Number of PDF-Files for peaking: " + np.ToString());

                if (np > 0)
                {
                    ShowLogMsg("");
                    int count = 1;
                    //List<string> peakList = new List<string>();
                    foreach (var file in toRead)
                    {
                        try
                        {
                            //await Task.Run(() => peakList.Add(DumpPeakDicomTags(file) + "_#"));
                            progressBar.Value = count * 100 / np;
                            ProgressTextBlock.Text = "Reading PDF-Files....";
                            count++;
                        }
                        catch
                        {
                            ShowLogMsg("Error reading " + file);
                            count++;
                        }
                    }                    

                    progressBar.Value = 0;
                    ProgressTextBlock.Text = "PDF-Peak finished - Press Start to begin Conversion";
                }
                else
                {
                    //ShowLogMsg("No PDF-Peak possible. Folder does not contain PDF-Files.");
                }
            }
            else
            {
                ShowLogMsg("\nNo PDF-Peak possible. Folder does not exist.");
            }
        }


        private void AboutMe_Click(object sender, RoutedEventArgs e)
        {
            FlexibleMessageBox.Show("See more of my apps on GitHub:\n\nhttps://github.com/Kiragroh\n\nPersonal information can be found on my LinkedIn:\n\nhttps://www.linkedin.com/in/maximilian-grohmann-b70588b1\n\nHave fun.\nMax", "About me");
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Disclaimer_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("This program is not tested for commercial or clinical use.\n\nYou use it at your own risk, and you are responsible for the interpretation of any results.", "Disclaimer");
        }

        

        private void selectDevFileButton_Click(object sender, RoutedEventArgs e)
        {
           
        }

        private void DevButton_Click(object sender, RoutedEventArgs e)
        {
            
           

        }
        
    }

    class LineUsingCoordinates
    {
        public static List<List<string>> getLineText(string path, int page, float[] coord)
        {
            //Create an instance of our strategy
            var t = new MyLocationTextExtractionStrategy();

            //Parse page 1 of the document above
            using (var r = new PdfReader(path))
            {
                for (var i = 0; i < r.NumberOfPages; i++)
                {
                    //var ex = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(r, 2, t);
                }
                // Calling this function adds all the chunks with their coordinates to the 
                // 'myPoints' variable of 'MyLocationTextExtractionStrategy' Class
                var ex = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(r, page, t);
            }
            // List of columns in one line
            List<string> lineWord = new List<string>();
            // temporary list for working around appending the <List<List<string>>
            List<string> tempWord;
            // List of rows. rows are list of string
            List<List<string>> lineText = new List<List<string>>();
            // List consisting list of chunks related to each line
            List<List<RectAndText>> lineChunksList = new List<List<RectAndText>>();
            //List consisting the chunks for whole page;
            List<RectAndText> chunksList;
            // List consisting the list of Bottom coord of the lines present in the page 
            List<float> bottomPointList = new List<float>();

            //Getting List of Coordinates of Lines in the page no matter it's a table or not
            foreach (var i in t.myPoints)
            {
                // If the coords passed to the function is not null then process the part in the 
                // given coords of the page otherwise process the whole page
                if (coord != null)
                {
                    if (i.Rect.Left >= coord[0] &&
                        i.Rect.Bottom >= coord[1] &&
                        i.Rect.Right <= coord[2] &&
                        i.Rect.Top <= coord[3])
                    {
                        float bottom = i.Rect.Bottom;
                        if (bottomPointList.Count == 0)
                        {
                            bottomPointList.Add(bottom);
                        }
                        else if (Math.Abs(bottomPointList.Last() - bottom) > 3)
                        {
                            bottomPointList.Add(bottom);
                        }
                    }
                }
                // else process the whole page
                else
                {
                    float bottom = i.Rect.Bottom;
                    if (bottomPointList.Count == 0)
                    {
                        bottomPointList.Add(bottom);
                    }
                    else if (Math.Abs(bottomPointList.Last() - bottom) > 3)
                    {
                        bottomPointList.Add(bottom);
                    }
                }
            }

            // Sometimes the above List will be having some elements which are from the same line but are
            // having different coordinates due to some characters like " ",".",etc.
            // And these coordinates will be having the difference of at most 4 points between 
            // their bottom coordinates. 

            //so to remove those elements we create two new lists which we need to remove from the original list 

            //This list will be having the elements which are having different but a little difference in coordinates 
            List<float> removeList = new List<float>();
            // This list is having the elements which are having the same coordinates
            List<float> sameList = new List<float>();

            // Here we are adding the elements in those two lists to remove the elements
            // from the original list later
            for (var i = 0; i < bottomPointList.Count; i++)
            {
                var basePoint = bottomPointList[i];
                for (var j = i + 1; j < bottomPointList.Count; j++)
                {
                    var comparePoint = bottomPointList[j];
                    //here we are getting the elements with same coordinates
                    if (Math.Abs(comparePoint - basePoint) == 0)
                    {
                        sameList.Add(comparePoint);
                    }
                    // here ae are getting the elements which are having different but the diference
                    // of less than 4 points
                    else if (Math.Abs(comparePoint - basePoint) < 4)
                    {
                        removeList.Add(comparePoint);
                    }
                }
            }

            // Here we are removing the matching elements of remove list from the original list 
            bottomPointList = bottomPointList.Where(item => !removeList.Contains(item)).ToList();

            //Here we are removing the first matching element of same list from the original list
            foreach (var r in sameList)
            {
                //bottomPointList.Remove(r);
            }

            // Here we are getting the characters of the same line in a List 'chunkList'.
            foreach (var bottomPoint in bottomPointList)
            {
                chunksList = new List<RectAndText>();
                for (int i = 0; i < t.myPoints.Count; i++)
                {
                    // If the character is having same bottom coord then add it to chunkList
                    if (bottomPoint == t.myPoints[i].Rect.Bottom)
                    {
                        chunksList.Add(t.myPoints[i]);
                    }
                    // If character is having a difference of less than 3 in the bottom coord then also
                    // add it to chunkList because the coord of the next line will differ at least 10 points
                    // from the coord of current line
                    else if (Math.Abs(t.myPoints[i].Rect.Bottom - bottomPoint) < 3)
                    {
                        chunksList.Add(t.myPoints[i]);
                    }
                }
                // Here we are adding the chunkList related to each line
                lineChunksList.Add(chunksList);
            }
            //bool sameLine = false;

            //Here we are looping through the lines consisting the chunks related to each line 
            foreach (var linechunk in lineChunksList)
            {
                var text = "";
                // Here we are looping through the chunks of the specific line to put the texts
                // that are having a cord jump in their left coordinates.
                // because only the line having table will be having the coord jumps in their 
                // left coord not the line having texts
                for (var i = 0; i < linechunk.Count - 1; i++)
                {
                    // If the coord is having a jump of less than 3 points then it will be in the same
                    // column otherwise the next chunk belongs to different column
                    if (Math.Abs(linechunk[i].Rect.Right - linechunk[i + 1].Rect.Left) < 3)
                    {
                        if (i == linechunk.Count - 2)
                        {
                            text += linechunk[i].Text + linechunk[i + 1].Text;
                        }
                        else
                        {
                            text += linechunk[i].Text;
                        }
                    }
                    else
                    {
                        if (i == linechunk.Count - 2)
                        {
                            // add the text to the column and set the value of next column to ""
                            text += linechunk[i].Text;
                            // this is the list of columns in other word its the row
                            lineWord.Add(text);
                            text = "";
                            text += linechunk[i + 1].Text;
                            lineWord.Add(text);
                            text = "";
                        }
                        else
                        {
                            text += linechunk[i].Text;
                            lineWord.Add(text);
                            text = "";
                        }
                    }
                }
                if (text.Trim() != "")
                {
                    lineWord.Add(text);
                }
                // creating a temporary list of strings for the List<List<string>> manipulation
                tempWord = new List<string>();
                tempWord.AddRange(lineWord);
                // "lineText" is the type of List<List<string>>
                // this is our list of rows. and rows are List of strings
                // here we are adding the row to the list of rows
                lineText.Add(tempWord);
                lineWord.Clear();
            }

            return lineText;
        }
    }
}
