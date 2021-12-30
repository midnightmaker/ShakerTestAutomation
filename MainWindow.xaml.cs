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
using System.Windows.Automation;
using System.Diagnostics; // for Process
using System.Threading; // for Thread.Sleep()
using System.Windows.Threading;
using System.IO;
using System.Windows.Automation.Text;
using System.Collections;
using System.Runtime.CompilerServices;

namespace ShakerTestAutomation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        AutomationElement aePhxForm = null;
        AutomationElement aeOmegaForm = null;
        AutomationElement aePhxToolbar = null;
        AutomationElement aeOmegaTestScriptTabControl = null;

        string[] ComboBoxItems = { "Linear X", "Linear X", "Linear X", "Linear X", "Orbital", "Orbital", "Orbital", "Orbital", "Orbital", "Orbital",
                                    "Double Orbital", "Double Orbital", "Double Orbital", "Double Orbital", "Double Orbital" };
        string[] TestSpeeds = {"100","300","500","700","100","300","500","700","900","1100","300","500","700","900","1100"};

        DateTime dtTestStart;
        DispatcherTimer timer;
        bool phxRecordingActive = false;
        int testCount = 0;
        string plateSerialNumber = "";
        public MainWindow()
        {
            InitializeComponent();
            timer = new DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 1);
            timer.Tick += Timer_Tick;
        }

        private void StartTimer()
        {
            dtTestStart = DateTime.Now;
            phxRecordingActive = false;
            timer.Start();
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            TimeSpan ts = DateTime.Now - dtTestStart;
            lblTestProgress.Content = plateSerialNumber + " " + ts.ToString();
            if( !phxRecordingActive && ts.TotalSeconds > 12 )
            {
                // And the test should now be running - wait 12 seconds for Omega to start shaking
                // Now start the Phosentix monitoring
                AutomationElement shakerTestStartDlg = GetShakerTestStartDialog();
                if (shakerTestStartDlg != null)
                {
                    AutomationElement comboBox = shakerTestStartDlg.FindFirst(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.AutomationIdProperty, "cmbShakingMode"));
                    SetSelectedComboBoxItem(comboBox, ComboBoxItems[testCount]);

                    SetSpeedTextBox(shakerTestStartDlg, TestSpeeds[testCount]);
                    
                    PressButton(shakerTestStartDlg, "StartButton");
                    
                    phxRecordingActive = true;
                    lblStatus.Content = ComboBoxItems[testCount] + " " + TestSpeeds[testCount] + "RPM";
                    testCount++;
                }

            }

            if ( ts.TotalSeconds > 50  && phxRecordingActive )
            {
                // Check to see if Omega dialog has popped up - this dialog will pop up after end of the current method file
                // and will be after the Phosentix results are displayed.
                AutomationElement okButton = GetOmegaInfoDialogPopupOkButton();
                if( okButton != null )
                {
                    // Close down the result widow on PhosentixShaker Window
                    timer.Stop();
                    AutomationElement reportViewer = aePhxForm.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Phosentix Insight Report Viewer"));

                    // Record the final values for this run
                    // NOTE: While this works, screen scraping the values takes a very long time to complete. Not sure why. The Phosentix
                    // insight applicaiton 1.3.5.0 has been updated to write the CSV file  instead of using the screen scraping method.
                    // 
                    //RecordFinalValues(aePhxForm, ComboBoxItems[testCount - 1], TestSpeeds[testCount - 1]);

                    WindowPattern windowPattern = reportViewer.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                    CloseWindow(windowPattern);

                    // Now Press Omega OK button in the script window messagebox
                    PressOKOnOMegaInfo( okButton );
                    if (testCount < 15)
                    {
                        PressButton(aePhxToolbar, "btnShakerTestControl");
                        StartTimer();
                        

                    }
                    else
                    {
                        MessageBox.Show("All tests complete");
                    }
                }
                
            }
        }

        private void RecordFinalValues(AutomationElement phxForm, string testType, string targetRPM)
        {
            // TextBoxes and Labels have to be read using separate functions due to the 
            // internal differences within the objects
            ArrayList values = new ArrayList();
            DateTime dt = DateTime.Now;

            values.Add(dt.ToString("G")); // Use general date format code
            
            // To keep excel from dropping leading zeros, add the apostrophe
            values.Add( "'" + plateSerialNumber );

            // Test Type is Orbital, Linear etc
            values.Add(testType);

            values.Add(GetLabelValue(phxForm, "CurrentRPMX"));
            values.Add(GetLabelValue(phxForm, "CurrentRPMY"));
            values.Add(GetLabelValue(phxForm, "AverageRPMX"));
            values.Add(GetLabelValue(phxForm, "AverageRPMY"));

            // And textboxes need special treatment.
            //values.Add(GetTextBoxText(phxForm, "TargetRPM"));
            values.Add(targetRPM);

            values.Add(GetLabelValue(phxForm, "TemperatureTL"));
            values.Add(GetLabelValue(phxForm, "TemperatureTR"));
            values.Add(GetLabelValue(phxForm, "TemperatureBL"));
            values.Add(GetLabelValue(phxForm, "TemperatureBR"));
            values.Add(GetLabelValue(phxForm, "TemperatureAvg"));

            StringBuilder sb = new StringBuilder();
            foreach( string s in values)
            {
                sb.Append(s);
                sb.Append(",");
            }

            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            string fileName = filePath + "\\" + plateSerialNumber + "-results.csv";

            bool addHeader = !File.Exists(fileName);
            using (StreamWriter outputFile = new StreamWriter(fileName, append: true))
            {
                if(addHeader)
                {
                    outputFile.WriteLine("Date,SerialNumber,Mode,CurrentRPMX,CurrentRPMY,AverageRPMX,AverageRPMY,TargetRPM,TemperatureTL,TemperatureTR,TemperatureBL,TemperatureBR,TemperatureAvg");
                }
                outputFile.WriteLine(sb.ToString().TrimEnd(','));
            }

        }

        private string GetLabelValue( AutomationElement phxForm, string automationID )
        {
            AutomationElement labelField = phxForm.FindFirst(TreeScope.Descendants,
                     new PropertyCondition(AutomationElement.AutomationIdProperty, automationID));

            if(labelField != null )
            {
                string text = labelField.Current.Name;
                text = text.Trim();
                text = text.TrimEnd('C');
                text = text.TrimEnd('°');
                text = text.Trim();
                return text;
            }

            return "Not found - " + automationID;

        }

        void SetSpeedTextBox( AutomationElement shakerTestStartDlg, string speed )
        {
            AutomationElement speedText = shakerTestStartDlg.FindFirst(TreeScope.Descendants,
                      new PropertyCondition(AutomationElement.AutomationIdProperty, "TargetSpeed"));

            ValuePattern vpTextBox1 = (ValuePattern)speedText.GetCurrentPattern(ValuePattern.Pattern);
            vpTextBox1.SetValue(speed);

        }

        string GetTextBoxText( AutomationElement form, string automationId)
        {
            AutomationElement textBox = form.FindFirst(TreeScope.Descendants,
                      new PropertyCondition(AutomationElement.AutomationIdProperty, automationId));

            if( textBox != null )
            {
                // Get required control patterns
                TextPattern targetTextPattern = textBox.GetCurrentPattern(TextPattern.Pattern) as TextPattern;
                if (targetTextPattern != null)
                {
                    TextPatternRange textRange = targetTextPattern.DocumentRange;
                    return textRange.GetText(20);
                }
                else
                {
                    return "TextPattern not found - " + automationId;
                }
            }

            return "Not Found - " + automationId;
        }

        private void CloseWindow(WindowPattern windowPattern)
        {
            try
            {
                windowPattern.Close();
            }
            catch (InvalidOperationException)
            {
                // object is not able to perform the requested action
                return;
            }
        }

        private void btnStartAutomation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AutomationElement aeDesktop = AutomationElement.RootElement;

                int numWaits = 0;
                if(aePhxForm == null)
                {
                    do
                    {
                        aePhxForm = aeDesktop.FindFirst(TreeScope.Children,
                          new PropertyCondition(AutomationElement.NameProperty, "Phosentix Insight"));
                        ++numWaits;
                        Thread.Sleep(100);
                    } while (aePhxForm == null && numWaits < 50);
                    if (aePhxForm == null)
                        throw new Exception("Failed to find Phosentix Insight");
                    else
                    {

                        lblStatus.Content = "Found Phosentix Insight Window...";
                        plateSerialNumber = GetTextBoxText(aePhxForm, "PlateSerialNumber");
                        if (plateSerialNumber.ToUpper().Contains("NOT"))
                        {
                            plateSerialNumber = "";
                            throw new Exception("Could not get serial number for MQC plate");
                        }
                    }
                    aePhxToolbar = aePhxForm.FindFirst(TreeScope.Children,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar));

                }



                //AutomationElement testDialog = this.GetShakerTestStartDialog();
                //SetSpeedTextBox(testDialog, "234");
                //AutomationElement comboBox = testDialog.FindFirst(TreeScope.Descendants,
                //       new PropertyCondition(AutomationElement.AutomationIdProperty, "cmbShakingMode"));
                //SetSelectedComboBoxItem(comboBox, ComboBoxItems[testCount]);

                //PressButton(testDialog, "StartButton");
                if( aeOmegaForm == null )
                {
                    do
                    {
                        AutomationElementCollection allWindows = aeDesktop.FindAll(TreeScope.Children,
                         new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                        // If this fails in future, be aware that the window Title has a space in it " Omega" not "Omega"
                        // SOmeone may change this at BMG in future in which case you will need to adjust the nameproperty below
                        aeOmegaForm = aeDesktop.FindFirst(TreeScope.Children,
                          new PropertyCondition(AutomationElement.NameProperty, " Omega"));

                        ++numWaits;
                        Thread.Sleep(100);
                    } while (aeOmegaForm == null && numWaits < 50);
                    if (aeOmegaForm == null)
                        throw new Exception("Failed to find Omega window");
                    else
                        lblStatus.Content += " Found OMEGA Window...";

                }


                AutomationElementCollection allChildren = aeOmegaForm.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                lblStatus.Content += $"Found {allChildren.Count } children";
                foreach (AutomationElement ae in allChildren)
                {
                    if (ae.Current.Name.Contains("Script"))
                    {
                        lblStatus.Content = "Found script window";
                        AutomationElement panel = ae.FindFirst(TreeScope.Children,
                            new PropertyCondition(AutomationElement.ClassNameProperty, "TPanel"));

                        if (panel != null)
                        {
                            AutomationElement tabControl = panel.FindFirst(TreeScope.Children,
                                          new PropertyCondition(AutomationElement.ClassNameProperty, "TcxTabControl"));

                            if (tabControl != null)
                            {
                                aeOmegaTestScriptTabControl = tabControl;

                                AutomationElementCollection allButtons = tabControl.FindAll(TreeScope.Children,
                                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
                                bool foundStart = false;
                                foreach (AutomationElement btn in allButtons)
                                {
                                    if (btn.Current.Name.Contains("Start"))
                                    {
                                        lblStatus.Content += "  Found start button";
                                        StartTest();
                                        foundStart = true;
                                        break;
                                    }
                                }

                                if( !foundStart )
                                {
                                    lblStatus.Content = "Unable to find OMEGA script start button";
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lblStatus.Content = ex.Message;
            }
        }

        void StartTest()
        {
            testCount = 0;
            //Press Shaker Test Start button on Phosentix window
            AutomationElement aeButton = aePhxToolbar.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "btnShakerTestControl"));

            InvokePattern btnPattern = aeButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            btnPattern.Invoke();

            //PressButton Omega script start button
            AutomationElementCollection allButtons = aeOmegaTestScriptTabControl.FindAll(TreeScope.Children,
                                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            foreach (AutomationElement btn in allButtons)
            {
                if (btn.Current.Name.Contains("Start"))
                {
                    InvokePattern btnPbtnStartattern = btn.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    btnPbtnStartattern.Invoke();
                    break;
                }
            }

            // Wait for first popup information dialog telling user to connect the plate via Bluetooth
            PressOKOnOMegaInfo();
            Thread.Sleep(100);
            // Now the second telling user to setup the linear 100RPM test
            PressOKOnOMegaInfo();

            lblTestProgress.Content = plateSerialNumber;
            StartTimer();


        }

        AutomationElement GetShakerTestStartDialog()
        {
            int waitCount = 0;
            do
            {
                AutomationElement aePhxShakerTestStartDlg = aePhxForm.FindFirst(TreeScope.Children,
                      new PropertyCondition(AutomationElement.NameProperty, "Shaker Test Start"));


                if (aePhxShakerTestStartDlg == null)
                    Thread.Sleep(100);
                else
                    return aePhxShakerTestStartDlg;
            }
            while (waitCount++ < 50);


            return null; ;

        }

        void PressButton(AutomationElement aeButtonContainer, string automationID)
        {

            AutomationElementCollection aeAllButtons = aeButtonContainer.FindAll(TreeScope.Children,
             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));


            AutomationElement aeButton = aeButtonContainer.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, automationID)); // "btnConnect"));

            InvokePattern btnPattern =
              aeButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            btnPattern.Invoke();
        }

        AutomationElement GetOmegaInfoDialogPopupOkButton()
        {
          AutomationElement msgBox = aeOmegaForm.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty, "TMsgBox"));
          if (msgBox != null)
          {
            AutomationElement okButton = msgBox.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "TButton"));
            return okButton;
          }
          return null;

        }
        bool PressOKOnOMegaInfo( AutomationElement button = null)
        {
            if( button != null )
            {
                InvokePattern btnPattern = button.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                btnPattern.Invoke();
                return true;
            }
            else
            {
                int waitCount = 0;
                do
                {
                    AutomationElement okButton = GetOmegaInfoDialogPopupOkButton();
                    if (okButton != null)
                    {
                        InvokePattern btnPattern = okButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        btnPattern.Invoke();
                        return true;
                    }
                    Thread.Sleep(100);
                }
                while (waitCount++ < 70);

                return false;

            }
        }

        public static void SetSelectedComboBoxItem(AutomationElement comboBox, string item)
        {
            AutomationPattern automationPatternFromElement = GetSpecifiedPattern(comboBox, "ExpandCollapsePatternIdentifiers.Pattern");

            ExpandCollapsePattern expandCollapsePattern = comboBox.GetCurrentPattern(automationPatternFromElement) as ExpandCollapsePattern;

            expandCollapsePattern.Expand();
            expandCollapsePattern.Collapse();

            AutomationElement listItem = comboBox.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, item));

            automationPatternFromElement = GetSpecifiedPattern(listItem, "SelectionItemPatternIdentifiers.Pattern");

            SelectionItemPattern selectionItemPattern = listItem.GetCurrentPattern(automationPatternFromElement) as SelectionItemPattern;

            selectionItemPattern.Select();
        }

        private static AutomationPattern GetSpecifiedPattern(AutomationElement element, string patternName)
        {
            AutomationPattern[] supportedPattern = element.GetSupportedPatterns();

            foreach (AutomationPattern pattern in supportedPattern)
            {
                if (pattern.ProgrammaticName == patternName)
                    return pattern;
            }

            return null;
        }

        ///--------------------------------------------------------------------

       

    }
}
