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
            timer.Start();
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            TimeSpan ts = DateTime.Now - dtTestStart;
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
                    testCount++;
                }

            }
            if ( ts.TotalSeconds > 90 )
            {
                // Close down the result widow on PhosentixShaker Window
                timer.Stop();
                AutomationElement reportViewer = aePhxForm.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Phosentix Insight Report Viewer"));

                WindowPattern windowPattern = reportViewer.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                CloseWindow(windowPattern);
                PressButton(aePhxToolbar, "btnShakerTestControl");
                // Now Press Omega OK button in the script window messagebox
                PressOKOnOMegaInfo();
                if(testCount < 15)
                {
                    StartTimer();
                    phxRecordingActive = false;
                    
                }
                else
                {
                    MessageBox.Show("All tests complete");
                }

            }
        }

        void SetSpeedTextBox( AutomationElement shakerTestStartDlg, string speed )
        {
            AutomationElement speedText = shakerTestStartDlg.FindFirst(TreeScope.Descendants,
                      new PropertyCondition(AutomationElement.AutomationIdProperty, "TargetSpeed"));

            ValuePattern vpTextBox1 = (ValuePattern)speedText.GetCurrentPattern(ValuePattern.Pattern);
            vpTextBox1.SetValue(speed);

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
                    lblStatus.Content = "Found Phosentix Insight Window...";

                aePhxToolbar = aePhxForm.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar));

                //AutomationElement testDialog = this.GetShakerTestStartDialog();
                //SetSpeedTextBox(testDialog, "234");
                //AutomationElement comboBox = testDialog.FindFirst(TreeScope.Descendants,
                //       new PropertyCondition(AutomationElement.AutomationIdProperty, "cmbShakingMode"));
                //SetSelectedComboBoxItem(comboBox, ComboBoxItems[testCount]);

                //PressButton(testDialog, "StartButton");

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

                AutomationElementCollection allChildren = aeOmegaForm.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

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
                                foreach (AutomationElement btn in allButtons)
                                {
                                    if (btn.Current.Name.Contains("Start"))
                                    {
                                        lblStatus.Content += "  Found start button";
                                        StartTest();         
                                    }
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
            //Press Shaker Test Start button
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
                }
            }

            // Wait for first popup information dialog telling user to connect the plate via Bluetooth
            PressOKOnOMegaInfo();
            // Now the second telling user to setup the linear 100RPM test
            PressOKOnOMegaInfo();

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

        void WaitForShakerTestStartDialogToShow()
        {
            AutomationElement aeToolbar = aePhxForm.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar));
        }

        bool PressOKOnOMegaInfo()
        {
            int waitCount = 0;
            do
            {
                AutomationElement msgBox = aeOmegaForm.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty, "TMsgBox"));
                if (msgBox != null)
                {
                    AutomationElement okButton = msgBox.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "TButton"));
                    if (okButton != null)
                    {
                        InvokePattern btnPattern = okButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        btnPattern.Invoke();
                        return true;
                    }
                }
                Thread.Sleep(100);
            }
            while (waitCount++ < 50);

            return false;
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
