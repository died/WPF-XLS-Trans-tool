using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xls_Trans_Tool_Wpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public DataTable Dt = new DataTable();

        public MainWindow()
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(UnhandledHandler);

            SetCulture();
            InitializeComponent();
            SetSaveEnable(false);
            ReadSetting();

            try
            {
                var result = typeof(Excel.Application);
                ShowStatusText($".Net:{Environment.Version} Excel:{result.ToString() == "Microsoft.Office.Interop.Excel.Application"}");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        }

        /// <summary>
        /// read option setting
        /// </summary>
        private void ReadSetting()
        {
            Title = $"{Title} {Assembly.GetExecutingAssembly().GetName().Version}";
            LineHeight.ItemsSource = new List<int> {1, 2, 3, 4, 5};
            //LineHeight.SelectedIndex = 0;

            RemarkCheckBox.IsChecked = Properties.Settings.Default.Remark;
            ShippingInstructionCheckBox.IsChecked = Properties.Settings.Default.ShippingInstruction;
            PriorityOrderTypeCheckBox.IsChecked = Properties.Settings.Default.PriorityOrderType;
            ActualManufacturerCheckBox.IsChecked = Properties.Settings.Default.Manufacturer;
            SupplierMaterialCheckBox.IsChecked = Properties.Settings.Default.Material;
            IssueDateCheckBox.IsChecked = Properties.Settings.Default.IssueDate;
            Additional1CheckBox.IsChecked = Properties.Settings.Default.Opt1;
            Additional2CheckBox.IsChecked = Properties.Settings.Default.Opt2;
            Additional3CheckBox.IsChecked = Properties.Settings.Default.Opt3;
            Additional4CheckBox.IsChecked = Properties.Settings.Default.Opt4;
            Additional5CheckBox.IsChecked = Properties.Settings.Default.Opt5;
            LineHeight.SelectedIndex = Properties.Settings.Default.LineHeight > 0 ? Properties.Settings.Default.LineHeight - 1 : 0;
        }

        /// <summary>
        /// save option setting
        /// </summary>
        /// <param name="shipping"></param>
        /// <param name="priority"></param>
        /// <param name="remark"></param>
        /// <param name="manufacturer"></param>
        /// <param name="material"></param>
        /// <param name="opt"></param>
        /// <param name="lineHeight"></param>
        private void SaveSetting(bool shipping, bool priority, bool remark, bool manufacturer, bool material,bool issueDate, bool[] opt, int lineHeight)
        {
            Properties.Settings.Default.Remark = remark;
            Properties.Settings.Default.ShippingInstruction = shipping;
            Properties.Settings.Default.PriorityOrderType = priority;
            Properties.Settings.Default.Manufacturer = manufacturer;
            Properties.Settings.Default.Material = material;
            Properties.Settings.Default.IssueDate = issueDate;
            Properties.Settings.Default.Opt1 = opt[0];
            Properties.Settings.Default.Opt2 = opt[1];
            Properties.Settings.Default.Opt3 = opt[2];
            Properties.Settings.Default.Opt4 = opt[3];
            Properties.Settings.Default.Opt5 = opt[4];
            Properties.Settings.Default.LineHeight = lineHeight;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// catch exception
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        static void UnhandledHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = (Exception)args.ExceptionObject;
            MessageBox.Show("UnhandledHandler caught : " + e.Message);
            if(e.InnerException!=null) MessageBox.Show("UnhandledHandler source : " + e.InnerException.Message);
            MessageBox.Show("UnhandledHandler trace : " + e.StackTrace);
        }

        /// <summary>
        /// For i18n
        /// </summary>
        private void SetCulture()
        {
            if (Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName.ToLower() == "zh")
            {
                Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("zh-TW");
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("zh-TW");
            }
        }

        #region interface
        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            SetProgressBar(0);
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                var result = dlg.ShowDialog(); // Show the dialog.
                if (result == true)
                {
                    string file = dlg.FileName;
                    SourceFileTextBox.Text = file;
                    try
                    {
                        ShowStatusText(Wording.LoadingFile);

                        ThreadPool.QueueUserWorkItem(o =>
                        {
                            var loadFile = ExcelTool.GetExcelFile(file);
                            if (loadFile.Success)
                            {
                                Dt = loadFile.Success ? (DataTable)loadFile.Data : null;
                                //Set progress bar if have data
                                if (Dt != null && Dt.Rows.Count > 0)
                                {
                                    SetupProgressBar(Dt.Rows.Count*2 + 2);
                                    SetSaveEnable(true);
                                    SetSaveLocation(file.Substring(0, file.LastIndexOf(".", StringComparison.Ordinal)) + "_" + DateTime.Now.ToFileTime() + ".xlsx");
                                    ShowStatusText(string.Format(Wording.LoadSuccessGotrows, Dt.Rows.Count));
                                }
                                else
                                {
                                    ShowStatusText(Wording.LoadFail);
                                }
                            }
                            else
                            {
                                ShowStatusText(loadFile.Message);
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        var msg = string.Format(Wording.OpenFileError, ex.Message);
                        ShowStatusText(msg);
                        Debug.WriteLine(msg);
                        MessageBox.Show(msg);
                    }
                }
            }
            catch (Exception exception)
            {
                ShowStatusText(exception.Message);
                Console.WriteLine(exception);
            }
        }

        /// <summary>
        /// Save and exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var location = SaveLocation;
            var sales = SalesTextBox.Text;
            var salesAssustant = SalesAssistantTextBox.Text;
            var shippingInstruction = ShippingInstructionCheckBox.IsChecked ?? false;
            var priorityOrderType = PriorityOrderTypeCheckBox.IsChecked ?? false;
            var remark = RemarkCheckBox.IsChecked ?? false;
            var manufacturer = ActualManufacturerCheckBox.IsChecked ?? false;
            var material = SupplierMaterialCheckBox.IsChecked ?? false;
            var issueDate = IssueDateCheckBox.IsChecked ?? false;
            var lineHeight = (int)LineHeight.SelectedValue;

            var addiitonalOptional = new[]
            {
                Additional1CheckBox.IsChecked ?? false, Additional2CheckBox.IsChecked ?? false,
                Additional3CheckBox.IsChecked ?? false, Additional4CheckBox.IsChecked ?? false,
                Additional5CheckBox.IsChecked ?? false
            };
            ThreadPool.QueueUserWorkItem(o =>
            {
                ShowStatusText(Wording.SaveingFile);
                SaveSetting(shippingInstruction, priorityOrderType, remark, manufacturer, material, issueDate, addiitonalOptional, lineHeight);
                SaveExcel(location, sales, salesAssustant, shippingInstruction, priorityOrderType, remark, manufacturer, material, issueDate, addiitonalOptional,lineHeight);
                ShowStatusText(Wording.SaveSuccess);
                Task.Delay(1500).ContinueWith(_ =>
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Environment.Exit(0);
                });
            });
        }
        #endregion

        /// <summary>
        /// doing save work
        /// </summary>
        /// <param name="location"></param>
        /// <param name="sales"></param>
        /// <param name="salesAssustant"></param>
        /// <param name="shipping"></param>
        /// <param name="priority"></param>
        /// <param name="remark"></param>
        /// <param name="manufacturer"></param>
        /// <param name="material"></param>
        /// <param name="issueDate"></param>
        /// <param name="ad">additional option 1~5</param>
        /// <param name="lineHeight"></param>
        private void SaveExcel(string location,string sales,string salesAssustant,bool shipping, bool priority, bool remark,bool manufacturer,bool material,bool issueDate, bool[] ad,int lineHeight)
        {
            var excelApp = new Excel.Application();//{Visible = true};
            var workbooks = excelApp.Workbooks;
            try
            {
                workbooks.Add();
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception Message: " + e.Message);
                if (e.InnerException != null) MessageBox.Show("InnerException Message: " + e.InnerException.Message);
                MessageBox.Show("Exception Trace : " + e.StackTrace);
                excelApp.Quit();
                return;
            }
            
            #region first sheet
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Name = "ForKeyIn";

            //add head
            for (int i = 0; i < Models.FuhsunHeaders.Length; i++)
            {
                workSheet.Cells[1, 1 + i] = Models.FuhsunHeaders[i];
            }

            #region find adidas data put into cell
            for (var i = 0; i < Dt.Rows.Count; i++)
            {
                //check empty row
                if(string.IsNullOrEmpty(Dt.Rows[i].Field<string>(1))|| string.IsNullOrEmpty(Dt.Rows[i].Field<string>(2))) continue;

                foreach (var m in Models.AdidasToFuhsunMapping)
                {
                    var value = Dt.Rows[i].Field<string>(ExcelTool.LettersToInt(m.Value) - 1);
                    //unit formatting
                    if (m.Value == "BP" && value == "YARD") value = "YRD";
                    //contact person
                    if (m.Value == "O")
                    {
                        var r = Regex.Match(value, @"[a-zA-Z0-9.]*");
                        if (r.Length > 0) value = r.Value;
                    }
                    //Material color to color code/desc
                    if (m.Value == "BM")
                    {
                        var ar = value.Split(new[] { ';', '/'}, StringSplitOptions.RemoveEmptyEntries);
                        var colorCode = new List<string>();
                        var colorDesc = new List<string>();
                        foreach (var t in ar)
                        {
                            var s = t.Trim();
                            if (s.Length > 5)
                            {
                                colorCode.Add(s.Substring(s.Length - 4, 4));
                                colorDesc.Add(s.Substring(0, s.Length - 4).Trim());
                            }
                            //only multi colorCode
                            else
                            {
                                colorCode.Add(t);
                            }
                        }
                        //try make sure the format is right
                        if (colorCode.Count > 0 && colorCode.Count == colorDesc.Count)
                        {
                            value = string.Join("/", colorCode);
                            workSheet.Cells[2 + i, "N"] = string.Join("/", colorDesc);
                        }
                        //only code
                        else
                        {
                            value = string.Join("/", colorCode);
                        }
                    }
                    workSheet.Cells[2 + i, m.Key] = value;
                }

                #region season
                var seasonColume = Dt.Rows[i].Field<string>(ExcelTool.LettersToInt("CE") - 1);
                if (!string.IsNullOrEmpty(seasonColume))
                {
                    var season = seasonColume.Split('-');
                    if (season.Length > 1)
                    {
                        workSheet.Cells[2 + i, "V"] = season[1];
                        workSheet.Cells[2 + i, "W"] = season[0];
                    }
                    else
                    {
                        workSheet.Cells[2 + i, "V"] = seasonColume;
                        workSheet.Cells[2 + i, "W"] = seasonColume;
                    }
                }
                #endregion

                #region contact person

                #endregion

                //TODO add custom input info
                //Sales
                workSheet.Cells[2 + i, "AU"] = sales;
                workSheet.Cells[2 + i, "AV"] = salesAssustant;

                StepProgressBar();
            }
            #endregion

            for (var j = 1; j <= Models.AdidasToFuhsunMapping.Count; j++)
            {
                ((Excel.Range)workSheet.Columns[j]).AutoFit();
            }
            StepProgressBar();
            #endregion

            #region 2nd sheet
            //2nd sheet
            Excel._Worksheet newWorksheet = (Excel.Worksheet)excelApp.Worksheets.Add();
            newWorksheet.Name = "Summary";

            #region make header/mapping
            var summaryHeader = Models.SummaryHeaders;
            var summaryMapping = Models.SummaryMapping;
            if (shipping)
            {
                summaryHeader.Add("Shipping\nInstruction");
                summaryMapping.Add("AW");
            }
            if (priority)
            {
                summaryHeader.Add("Priority &\nOrder Type");
                summaryMapping.Add("CF");
            }
            if (remark)
            {
                summaryHeader.Add("Remarks");
                summaryMapping.Add("CJ");
            }
            if (manufacturer)
            {
                summaryHeader.Add("Actual\nManufacturer");
                summaryMapping.Add("AC");
            }
            if (material)
            {
                summaryHeader.Add("Supplier\nMaterial");
                summaryMapping.Add("BA");
            }
            if (ad[0])
            {
                summaryHeader.Add("Additional\nOptional 1");
                summaryMapping.Add("CK");
            }
            if (ad[1])
            {
                summaryHeader.Add("Additional\nOptional 2");
                summaryMapping.Add("CL");
            }
            if (ad[2])
            {
                summaryHeader.Add("Additional\nOptional 3");
                summaryMapping.Add("CM");
            }
            if (ad[3])
            {
                summaryHeader.Add("Additional\nOptional 4");
                summaryMapping.Add("CN");
            }
            if (ad[4])
            {
                summaryHeader.Add("Additional\nOptional 5");
                summaryMapping.Add("CO");
            }

            //add issue Date 20181130
            if (issueDate)
            {
                summaryHeader.Insert(0, "Issue Date");
                summaryMapping.Insert(0, "A");
            }

            //add contact person
            summaryHeader.Add("Contact Person");
            summaryMapping.Add("O");

            //move remark to checkbox
            //summaryHeader.AddRange(ExcelTool.RemarkHeaders);
            //summaryMapping.AddRange(ExcelTool.RemarkMapping);
            #endregion

            //add header
            for (int i = 0; i < summaryHeader.Count; i++)
            {
                newWorksheet.Cells[1, 1 + i] = summaryHeader[i];
            }
            //add hidden column on summary
            var hiddenColumn = new List<int>();

            //fill content
            for (var i = 0; i < Dt.Rows.Count; i++)
            {
                if (string.IsNullOrEmpty(Dt.Rows[i].Field<string>(1)) || string.IsNullOrEmpty(Dt.Rows[i].Field<string>(2))) continue;

                for (var j = 0; j < summaryMapping.Count; j++)
                {
                    var source = summaryMapping[j];
                    if(string.IsNullOrEmpty(source)) continue;
                    var value = Dt.Rows[i].Field<string>(ExcelTool.LettersToInt(source) - 1);
                    //unit formatting
                    if (source == "BP" && value == "YARD") value = "YRD";
                    if (source == "CJ") value = Dt.Rows[i].Field<string>(ExcelTool.LettersToInt("AU") - 1) + value;
                    //contact user
                    if (source == "O")
                    {
                        value = Regex.Match(value, @"[a-zA-Z0-9.]*").Value;
                        hiddenColumn.Add(j+1);
                    }
                    //matrial color => color code/desc
                    if (source == "BM" && !string.IsNullOrEmpty(value))
                    {
                        var ar = value.Split(new[] { ';', '/' }, StringSplitOptions.RemoveEmptyEntries);
                        var colorCode = new List<string>();
                        var colorDesc = new List<string>();
                        foreach (string t in ar)
                        {
                            var s = t.Trim();
                            if (s.Length > 5)
                            {
                                colorCode.Add(s.Substring(s.Length - 4, 4));
                                colorDesc.Add(s.Substring(0, s.Length - 4).Trim());
                            }
                            //only multi colorCode
                            else
                            {
                                colorCode.Add(t);
                            }
                        }
                        //try make sure the format is right
                        if (colorCode.Count > 0 && colorCode.Count == colorDesc.Count)
                        {
                            value = string.Join("/", colorCode);
                            newWorksheet.Cells[2 + i, "F"] = string.Join("/", colorDesc);
                        }
                        //only code
                        else
                        {
                            value = string.Join("/", colorCode);
                        }
                    }
                    newWorksheet.Cells[2 + i, j+1] = value;
                }

                //for YTI batch field
                try
                {
                    var last = summaryHeader.Count;
                    //put batch column's value to last-1
                    var v = Dt.Rows[i].Field<string>(ExcelTool.LettersToInt("CQ") - 1);
                    if (!string.IsNullOrEmpty(v))
                    {
                        var range = newWorksheet.Cells[2 + i, last-1] as Excel.Range;
                        if (range != null && range.Value2 != null)
                        {
                            newWorksheet.Cells[2 + i, last-1] = range.Value2.ToString() + "," + v;
                        }
                        else
                        {
                            newWorksheet.Cells[2 + i, last-1] = v;
                        }
                    }   
                }
                catch (Exception)
                {
                    //do nothing
                    //Console.WriteLine(ex.Message);
                }

                StepProgressBar();
            }
            //fit columns
            for (var j = 1; j <= Models.SummaryHeaders.Count; j++)
            {
                ((Excel.Range)newWorksheet.Columns[j]).AutoFit();
            }

            //set summary row height
            for (var i = 2; i <= Dt.Rows.Count+1; i++)
            {

                ((Excel.Range)workSheet.Rows[i]).RowHeight = 16 * lineHeight;
                ((Excel.Range)newWorksheet.Rows[i]).RowHeight = 16 * lineHeight;
            }

            //hidden column
            for (int i = 0; i < hiddenColumn.Count; i++)
            {
                ((Excel.Range)newWorksheet.Columns[hiddenColumn[i]]).Hidden = true;
            }

            StepProgressBar();
            #endregion

            workSheet.SaveAs(location);

            //close
            //excelApp.Workbooks.Close();
            workbooks.Close();
            excelApp.Quit();
            //wait release
            while (Marshal.ReleaseComObject(excelApp) != 0) { }
            while (Marshal.ReleaseComObject(workbooks) != 0) { }
            //excelApp = null;
            //workbooks = null;
        }

        #region control
        public string SaveLocation
        {
            get => TargetFileTextBox.Text;
            set => TargetFileTextBox.Text = value;
        }

        /// <summary>
        /// Show Status Text
        /// </summary>
        /// <param name="text"></param>
        private void ShowStatusText(string text)
        {
            Dispatcher.BeginInvoke(new Action(() => { ToolStripStatus.Text = text; }));
        }
        #endregion

        #region ProgressBar control
        private void SetupProgressBar(int max)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ToolStripProgressBar.Maximum = max;
                ToolStripProgressBar.Minimum = 0;
                ToolStripProgressBar.Value = 0;
            }));
        }

        private void SetProgressBar(int val)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ToolStripProgressBar.Value = val;
            }));
        }

        private void StepProgressBar()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ToolStripProgressBar.Value++;
            }));
        }
        #endregion

        /// <summary>
        /// Save
        /// </summary>
        /// <param name="state"></param>
        private void SetSaveEnable(bool state)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SaveButton.IsEnabled = state;
                TargetFileTextBox.IsEnabled = state;
            }));
        }

        private void SetSaveLocation(string str)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SaveLocation = str;
            }));
        }

        /// <summary>
        /// Change font size when window state change
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnWindowStateChange(object sender, EventArgs e)
        {
            //ShowStatusText(this.WindowState.ToString());
            var buttonList = new List<Button>();
            var labelList = new List<Label>();
            var textBoxList = new List<TextBox>();
            var textBlockList = new List<TextBlock>();
            var checkBoxList = new List<CheckBox>();
            GetLogicalChildCollection(this, labelList);
            GetLogicalChildCollection(this, buttonList);
            GetLogicalChildCollection(this, textBoxList);
            GetLogicalChildCollection(this, textBlockList);
            GetLogicalChildCollection(this, checkBoxList);
            foreach (var v in labelList)
            {
                v.FontSize = WindowState == WindowState.Maximized ? 18 : 12;
            }
            foreach (var v in buttonList)
            {
                v.FontSize = WindowState == WindowState.Maximized ? 18 : 12;
            }
            foreach (var v in textBoxList)
            {
                v.FontSize = WindowState == WindowState.Maximized ? 18 : 12;
            }
            foreach (var v in textBlockList)
            {
                v.FontSize = WindowState == WindowState.Maximized ? 18 : 12;
            }
            foreach (var v in checkBoxList)
            {
                v.FontSize = WindowState == WindowState.Maximized ? 18 : 12;
            }
        }

        /// <summary>
        /// Get same type controller
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <param name="logicalCollection"></param>
        private static void GetLogicalChildCollection<T>(DependencyObject parent, List<T> logicalCollection) where T : DependencyObject
        {
            IEnumerable children = LogicalTreeHelper.GetChildren(parent);
            foreach (object child in children)
            {
                if (!(child is DependencyObject)) continue;
                DependencyObject depChild = child as DependencyObject;
                if (child is T)
                {
                    logicalCollection.Add(child as T);
                }
                GetLogicalChildCollection(depChild, logicalCollection);
            }
        }
    }
}
