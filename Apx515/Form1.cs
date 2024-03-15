#define xREAD_BD_ADDR_BEFORE_TEST

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FTD2XX_NET;
using System.IO;
using System.Reflection;
using AudioPrecision.API;
using System.Threading;
using System.Media;

namespace Apx515
{
    public partial class Form1 : Form
    {
        public APx500 APx;

        public IMeterGraph CurrentMeter;

        public string pathConfigFile;
        public string pathLogFile;

        public string gAppName;

        public string configVbatComport;
        public string targetSerialVbatComport;
        public string configKeyComport;
        public string targetSerialKeyComport;
        public string configVDDComport;          
        public string targetSerialVDDComport;

        public string configKeyDelay;
        int intconfigKeyDelay;

        public string gModelName = "";

        public bool gDebugMode = false;

        public string gTestLevel;
        public string gTestThd;
        public string gTestStepFrequencyResponse;

        public double gConfigLimitLowerLevel;
        public double gConfigLimitUpperLevel;
        public double gConfigLimitUpperThd;
        public double gConfigLimitLowerSnr;
        public double gConfigLimitUpperBalance;

        public double gMeasureValueLevelLeft;
        public double gMeasureValueLevelSub;
        public double gMeasureValueLevelRight;
        public double gMeasureValueSnrLeft;
        public double gMeasureValueSnrSub;
        public double gMeasureValueSnrRight;
        public double gMeasureValueThdLeft;
        public double gMeasureValueThdSub;
        public double gMeasureValueThdRight;
        public double gMeasureValueBalance;

        public int gCount = 0;

        static public string dutFullBdAddressAgent;

        static public string gSoundNg1 = System.IO.Directory.GetCurrentDirectory() + "\\" + "sound\\" + "ng_merge.wav";

        public struct PathAndMeasurement
        {
            public ISignalPath Path;
            public ISequenceMeasurement Meas;

            public PathAndMeasurement(ISignalPath A, ISequenceMeasurement B)
            {
                this.Path = A;
                this.Meas = B;
            }
        }
        List<PathAndMeasurement> gMeasurementList = new List<PathAndMeasurement>();

        static public int gCountFrequencyPoint = 0;
        public string[] gFrequencyPoint = new string[20];
        public string[] gConfigFrequencyLimitLower = new string[20];
        public string[] gConfigFrequencyLimitUpper = new string[20];
        public double[] gMeasureValueSweepLeft = new double[20];
        public double[] gMeasureValueSweepRight = new double[20];
        public double[] gMeasureValueSweepSub = new double[20];
        public bool[] gNgFreq = new bool[20];


        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            //close APx
            if (APx != null)
                APx.Exit();
        }

        #region SET_UART_PORT
        private bool setUartPort()
        {
            int flag_ng = 0;

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Allocate storage for device info list
            FTD2XX_NET.FTDI.FT_DEVICE_INFO_NODE[] ftdiDeviceList;

            UInt32 ftdiDeviceCount = 0;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.GetNumberOfDevices(ref ftdiDeviceCount);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }

                // If devices available
                if (ftdiDeviceCount != 0)
                {
                    // Allocate storage for device info list
                    ftdiDeviceList = new FTD2XX_NET.FTDI.FT_DEVICE_INFO_NODE[ftdiDeviceCount];

                    // Populate our device list
                    ftStatus = myFtdiDevice.GetDeviceList(ftdiDeviceList);

                    if (ftStatus != FTDI.FT_STATUS.FT_OK)
                    {
                        flag_ng = 1;
                    }

                    if (flag_ng != 1)
                    {
                        for (int i = 0; i < ftdiDeviceCount; i++)
                        {
                            ftStatus = myFtdiDevice.OpenBySerialNumber(ftdiDeviceList[i].SerialNumber);

                            if (ftStatus == FTDI.FT_STATUS.FT_OK)
                            {
                                string merong;
                                myFtdiDevice.GetCOMPort(out merong);

                                if (merong == configVbatComport)
                                {
                                    targetSerialVbatComport = ftdiDeviceList[i].SerialNumber;
                                }

                                if (merong == configKeyComport)
                                {
                                    targetSerialKeyComport = ftdiDeviceList[i].SerialNumber;
                                }
                            }

                            myFtdiDevice.Close();

                        }
                    }

                    if ((targetSerialKeyComport == "") || (targetSerialVbatComport == ""))
                    {
                        flag_ng = 1;
                    }
                }
                else
                {
                    flag_ng = 1;
                }
            }

            if (flag_ng == 1)
                return false;
            else
                return true;
        }
        #endregion //#region SET_UART_PORT

        private void Form1_Load(object sender, EventArgs e)
        {
            initDataGridView();
            initDataGridViewSweep();

            configVbatComport = "";
            targetSerialVbatComport = "";
            configKeyComport = "";
            targetSerialKeyComport = "";

            dutFullBdAddressAgent = "";

            loadConfig();

            this.Text = gModelName + " " + gAppName + " V" + this.ProductVersion + " (password)";

            if (setUartPort()) // check all UART jig for testing are connected first
            {
                // initialize uart gpio port
                vbat_off();
                key_off();
                Thread.Sleep(500);

                try
                {
                    // Connect to APx500
                    APx = new APx500();

                    // Make the APx500 application window visible
                    APx.Visible = gDebugMode;

                    // Open the project file
                    APx.OpenProject(System.IO.Directory.GetCurrentDirectory() + "\\" + "config\\" + gModelName + "_PBA_AUDIO_TEST.approjx");

                    getMeasurementName();
                    checkLogFile();

                }
                catch (Exception ex)
                {
                    // Display an error message
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    Close();
                }
                finally
                {
                }
            }
            else
            {
                MessageBox.Show("Check UART-Jig Connection & Config file!!!!");
                Close();
            }

            //getMeasurementName();
        }

        private void NgSound()
        {

            SoundPlayer sound1 = new SoundPlayer(gSoundNg1);
            sound1.Play();

        }
        private void openpassworddlg(string a)
        {

            passwordbox password = new passwordbox(a);
            password.ShowDialog();
            //clearwindow();

        }

        private void initDataGridView() // data grid view for measured data (except sweep)
        {
            dataGridView1.ColumnCount = 7;
            dataGridView1.ColumnHeadersVisible = true;

            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.Blue;
            columnHeaderStyle.Font = new Font("Arial", 8, FontStyle.Regular);
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 200;
            dataGridView1.Columns[4].Width = 200;
            dataGridView1.Columns[5].Width = 200;
            dataGridView1.Columns[6].Width = 200;

            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView1.Columns[0].Name = "No.";
            dataGridView1.Columns[1].Name = "Item";
            dataGridView1.Columns[2].Name = "Limit_Lower";
            dataGridView1.Columns[3].Name = "Value";
            dataGridView1.Columns[4].Name = "Limit_Upper";
            dataGridView1.Columns[5].Name = "Mic. Type";
            dataGridView1.Columns[6].Name = "Result";
        }

        private void initDataGridViewSweep() // data grid view for sweep measured data
        {
            dataGridView2.ColumnCount = 4;
            dataGridView2.ColumnHeadersVisible = true;

            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
            DataGridViewCellStyle columnCellStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.Blue;
            columnHeaderStyle.Font = new Font("Arial", 8, FontStyle.Regular);
            dataGridView2.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            columnCellStyle.Font = new Font("Arial", 6, FontStyle.Regular);
            dataGridView2.DefaultCellStyle = columnCellStyle;


            dataGridView2.Columns[0].Width = 50;
            dataGridView2.Columns[1].Width = 300;
            dataGridView2.Columns[2].Width = 200;
            dataGridView2.Columns[3].Width = 200;

            dataGridView2.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView2.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView2.Columns[0].Name = "No.";
            dataGridView2.Columns[1].Name = "Freq.";
            dataGridView2.Columns[2].Name = "main mic.";
            dataGridView2.Columns[3].Name = "sub mic.";
        }

        private void loadConfig()
        {
            // using ini style config
            pathConfigFile = System.IO.Directory.GetCurrentDirectory() + "\\" + "config\\" + "config.ini";
            string tempString;

            IniReadWrite IniReader = new IniReadWrite();

            tempString = IniReader.IniReadValue("CONFIG", "KEY_Delay", pathConfigFile);
            configKeyDelay = tempString;
            Int32.TryParse(tempString, out intconfigKeyDelay);

            tempString = IniReader.IniReadValue("CONFIG", "MODEL_NAME", pathConfigFile);
            gModelName = tempString;

            tempString = IniReader.IniReadValue("CONFIG", "COM_VBAT", pathConfigFile);
            configVbatComport = tempString;

            tempString = IniReader.IniReadValue("CONFIG", "COM_VDD", pathConfigFile);
            configVDDComport = tempString;

            tempString = IniReader.IniReadValue("CONFIG", "COM_KEY", pathConfigFile);
            configKeyComport = tempString;


            tempString = IniReader.IniReadValue("CONFIG", "APPNAME", pathConfigFile);
            gAppName = tempString;
            lb_AppName.Text = gModelName + " " + gAppName + " V" + this.ProductVersion;
            lb_AppName.BackColor = Color.LightYellow;
            lb_AppName.ForeColor = Color.Green;

            tempString = IniReader.IniReadValue("CONFIG", "DEBUG_MODE", pathConfigFile);
            if (tempString == "TRUE")
            {
                gDebugMode = true;
            }
            else if (tempString == "FALSE")
            {
                gDebugMode = false;
            }
            else
            {
                gDebugMode = false;
            }

            // setup measure frequency point (stepped frequency response)
            // can add test point, limit = 20
            string tempFreqPoint = "";
            for (int i = 0; i < 20; i++)
            {
                tempFreqPoint = String.Format("POINT_{0}", i.ToString("D2"));
                //updateStatus(String.Format("High Temp(): {0}", count.ToString()));
                tempString = IniReader.IniReadValue("SWEEP_POINT", tempFreqPoint, pathConfigFile);
                if (tempString == "END")
                {
                    gCountFrequencyPoint = i;
                    break;
                }
                else
                {
                    // separate
                    gFrequencyPoint[i] = tempString;
                }
            }

            // set sweep data
            if (gCountFrequencyPoint > 0)
            {
                for (int k = 0; k < gCountFrequencyPoint; k++)
                {
                    string[] rowInfoFreq = { k.ToString(), "Sweep_" + gFrequencyPoint[k] + "Hz", "-", "-" };
                    dataGridView2.Rows.Add(rowInfoFreq);
                }
            }
        }

        private void getMeasurementName()
        {
            int count = 0;
            foreach (ISignalPath signalPath in APx.Sequence)
            {
                if (signalPath.Checked)
                {
                    foreach (ISequenceMeasurement measurement in signalPath)
                    {
                        //count every checked measurement
                        if (measurement.Checked)
                        {
                            string abcd = measurement.Name;
                            if (measurement.Name != "Signal Path Setup") // exclude the measurement name of "Signal Path Setup"
                            {
                                gMeasurementList.Add(new PathAndMeasurement(signalPath, measurement));

                                count++;
                            }
                        }
                    }
                }
            }
        }

        private void clearScreen()
        {
            // clear measurement values
            gMeasureValueLevelLeft = 0.000;
            gMeasureValueLevelRight = 0.000;
            gMeasureValueLevelSub = 0.000;
            gMeasureValueSnrLeft = 0.000;
            gMeasureValueSnrRight = 0.000;
            gMeasureValueSnrSub = 0.000;
            gMeasureValueThdLeft = 0.000;
            gMeasureValueThdRight = 0.000;
            gMeasureValueThdSub = 0.000;
            gMeasureValueBalance = 0.000;

            gMeasureValueSweepLeft.Initialize();
            gMeasureValueSweepRight.Initialize();
            gMeasureValueSweepSub.Initialize();

            // clear screen
            if (pictureBox1.Image != null)
            {
                pictureBox1.Image.Dispose();
                pictureBox1.Image = null;
            }

            if (pictureBox2.Image != null)
            {
                pictureBox2.Image.Dispose();
                pictureBox2.Image = null;
            }

            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();

            for (int i = 0; i < gCountFrequencyPoint; i++)
            {
                dataGridView2.Rows[i].Cells[2].Value = "";
                dataGridView2.Rows[i].Cells[3].Value = "";
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            bool flag_ng = false;

            string current_seq_name = "";

            btnStart.Enabled = false;

            clearScreen();
            lb_Result.Text = "...TESTING...";
            lb_Result.BackColor = Color.LightYellow;
            Application.DoEvents();

            vbat_on();
            Delay(1000);

            key_on();    
            Delay(intconfigKeyDelay);
            //Delay(2500);

            key_off();
            Delay(1000);

            sendMainMicLoopbackCmd();
            Delay(3000);

            gCount = 0;
            foreach (PathAndMeasurement seq in gMeasurementList)
            {
                try
                {
                    lb_Result.Text = "...TESTING... " + "(" + seq.Meas.Name + ")";
                    if((current_seq_name != seq.Path.Name) && (current_seq_name != ""))
                    {
                        // swithing to sub microphone loopback
                        Console.WriteLine("Switching to Sub mic. execute Sub mic loopback mode");
                        sendSubMicLoopbackCmd();
                        Thread.Sleep(1000);
                    }

                    if (!RunMeasurement(seq.Path, seq.Meas, gCount))
                        flag_ng = true;

                    current_seq_name = seq.Path.Name;
                }
                catch 
                {
                    //
                }

                gCount++;
            }

#if (READ_BD_ADDR_BEFORE_TEST)
            openUartForBDCmd();
            Delay(500);
#endif

            // update log
            writeLog(flag_ng);

            if (flag_ng)
            {
                lb_Result.Text = "FAIL";
                lb_Result.BackColor = Color.Red;
                NgSound();
            }
            else
            {
                lb_Result.Text = "PASS";
                lb_Result.BackColor = Color.SkyBlue;
            }

            btnStart.Enabled = true;

            // turn off power
            vbat_off();
            key_off();
        }

        private int FindSweepPointIndex(double[] xValues, double XValue)
        {
            //allow for a small tolerance in case measured value is slightly off from desired value
            XValue *= 1.001;

            int closestIndex = 0;
            double lastValue = Double.MinValue;

            //iterate through each value in the list until we find the closest match without
            //going over the desired value
            for (int i = 0; i < xValues.Length; i++)
            {
                if (xValues[i] > lastValue && xValues[i] <= XValue)
                {
                    closestIndex = i;
                    lastValue = xValues[i];

                }
                else
                {
                    if (Math.Abs(XValue - lastValue) > Math.Abs(XValue - xValues[i]))
                    {
                        closestIndex = i;
                    }

                    break;
                }
            }

            return closestIndex;
        }

        private bool RunMeasurement(ISignalPath SignalPath, ISequenceMeasurement Meas, int count)
        {
            bool testResult = false;

            Application.DoEvents();

            //make the measurement active
            APx.ShowMeasurement(SignalPath.Name, Meas.Name);

            try
            {
                APx.Sequence[SignalPath.Name][Meas.Name].Run();

                // get the object which represents the Rms Level meter values
                ISequenceResult result;

                switch (Meas.Name)
                {
                    case "Level and Gain":
                        result = APx.Sequence[SignalPath.Name][Meas.Name].SequenceResults[MeasurementResultType.LevelMeter];
                        testResult = UpdateResult(count, result, Meas.Name, SignalPath.Name);

                        break;
                    case "THD+N":
                        result = APx.Sequence[SignalPath.Name][Meas.Name].SequenceResults[MeasurementResultType.ThdNRatioMeter];
                        testResult = UpdateResult(count, result, Meas.Name, SignalPath.Name);
                        break;
                    case "Signal to Noise Ratio":
                        result = APx.Sequence[SignalPath.Name][Meas.Name].SequenceResults[MeasurementResultType.SignalToNoiseRatioMeter];
                        testResult = UpdateResult(count, result, Meas.Name, SignalPath.Name);
                        break;

                    case "Stepped Frequency Sweep":
                        result = APx.Sequence[SignalPath.Name][Meas.Name].SequenceResults[MeasurementResultType.LevelVsFrequency];
                        testResult = UpdateResultSweep(count, result, Meas.Name, SignalPath.Name);
                        IGraph edd = (IGraph)APx.SteppedFrequencySweep.Level;
                        edd.CopyToClipboard();
                        if(SignalPath.Name == "main_mic")
                        {
                            pictureBox1.Image = Clipboard.GetImage();
                        }
                        else if(SignalPath.Name == "sub_mic")
                        {
                            pictureBox2.Image = Clipboard.GetImage();
                        }
                        break;

                    case "Frequency Response":
                        result = APx.Sequence[SignalPath.Name][Meas.Name].SequenceResults[MeasurementResultType.LevelVsFrequency];
                        testResult = UpdateResultSweep(count, result, Meas.Name, SignalPath.Name);
                        IGraph add = (IGraph)APx.FrequencyResponse.Level;
                        add.CopyToClipboard();
                        if(SignalPath.Name == "main_mic")
                        {
                            pictureBox1.Image = Clipboard.GetImage();
                        }
                        else if(SignalPath.Name == "sub_mic")
                        {
                            pictureBox2.Image = Clipboard.GetImage();
                        }
                        break;
                }
            }
            catch
            {

            }

            APx.Sequence[SignalPath.Name][Meas.Name].Show();

            Application.DoEvents();

            return testResult;
        }

        //Check to see if(a meter value passes limit checks
        private static bool PassedLimitCheck(double meterValue, double lowerLimitValue, double upperLimitvalue)
        {
            bool Passed = true;

            //When a limit value is undefined, it has a value of Double.NaN
            //Check to make sure the limit value is !Double.NaN, and) check to see that the
            //meter value complies with the limit
            if (!double.IsNaN(lowerLimitValue) && meterValue < lowerLimitValue)
                Passed = false;

            //When a limit value is undefined, it has a value of Double.NaN
            //Check to make sure the limit value is !Double.NaN, and) check to see that the
            //meter value complies with the limit
            if (!double.IsNaN(upperLimitvalue) && meterValue > upperLimitvalue)
                Passed = false;

            return Passed;
        }

        private bool UpdateResult(int count, ISequenceResult result, string measureName, string sigPathName)
        {
            /*
             * get all meter value from each measurement, and display it
             * save meter value to global variable to save log
             */
            bool AllPass = true;

            double[] meterValues = result.GetMeterValues();
            double[] limitsLower = result.GetMeterLowerLimitValues();
            double[] limitsUpper = result.GetMeterUpperLimitValues();

            if (measureName == "Level and Gain")
            {
                if (sigPathName == "main_mic")
                {
                    gMeasureValueLevelLeft = meterValues[0];
                    //gMeasureValueLevelRight = meterValues[1];
                }
                else if (sigPathName == "sub_mic")
                {
                    gMeasureValueLevelSub = meterValues[0];
                }
            }
            else if (measureName == "THD+N")
            {
                if (sigPathName == "main_mic")
                {
                    gMeasureValueThdLeft = meterValues[0];
                    //gMeasureValueThdRight = meterValues[1];
                }
                else if (sigPathName == "sub_mic")
                {
                    gMeasureValueThdSub = meterValues[0];
                }
            }
            else if (measureName == "Signal to Noise Ratio")
            {
                if(sigPathName == "main_mic")
                {
                    gMeasureValueSnrLeft = meterValues[0];
                    //gMeasureValueSnrRight = meterValues[1];
                }
                else if(sigPathName == "sub_mic")
                {
                    gMeasureValueSnrSub = meterValues[0];
                }
            }

            for (int i = 0; i < meterValues.Length; i++)
            {
                if (PassedLimitCheck(meterValues[i], limitsLower[i], limitsUpper[i]))
                {
                    //
                }
                else
                {
                    //
                    AllPass = false;
                }
            }

            // update result
            string[] rowInfo = { count.ToString(), measureName, limitsLower[0].ToString("F3"), meterValues[0].ToString("F3"), limitsUpper[0].ToString("F3"), sigPathName, "-" };
            dataGridView1.Rows.Add(rowInfo);
            if (AllPass)
            {
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[6].Value = "PASS";
                dataGridView1.Rows[count].DefaultCellStyle.BackColor = Color.SkyBlue;
            }
            else
            {
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[6].Value = "FAIL";
                dataGridView1.Rows[count].DefaultCellStyle.BackColor = Color.Red;
            }

            if (AllPass) return true;
            else return false;
        }

        private bool UpdateResultSweep(int count, ISequenceResult result, string measureName, string sigPathName)
        {
            /*
             * meter value cannot be get when measurement is Sweep (Stepped Frequency Sweep)
             * So, pass/fail result will be base on the limit configuration
             * the all data for frequency point (configured in project file) will get
             * and we will get the measured value of nearst frequency point in that (find designated frequency (config.ini) in all frequency)
             */
            bool AllPass = true;
            int nearest_freq = 0;

            // update sweep result
            double[] MeasureValueSweepLeft = new double[gCountFrequencyPoint];
            //double[] MeasureValueSweepRight = new double[gCountFrequencyPoint];

            double[] XValues_Left = result.GetXValues(0);
            double[] YValues_Left = result.GetYValues(0);

            //double[] XValues_Right = result.GetXValues(1);
            //double[] YValues_Right = result.GetYValues(1);

            byte sweepResultPosition = 2; // default is dataGridView2.Rows[i].Cells[2]

            bool a = result.PassedLowerLimitCheck;
            bool b = result.PassedUpperLimitCheck;
            bool c = result.HasErrorMessage;
            bool d = result.HasMeterValues;
            bool e = result.HasXYValues;

            string abcd = result.Name;

            /* string to compare will be variable Sequence name in APx project file */
            if(sigPathName == "main_mic") { sweepResultPosition = 2; }
            else if(sigPathName == "sub_mic") { sweepResultPosition = 3; }

            if (!(result.PassedLowerLimitCheck && result.PassedUpperLimitCheck && !result.HasErrorMessage))
            {
                AllPass = false;
            }

            // update result
            string[] rowInfo = { count.ToString(), measureName, "-", "-", "-", sigPathName, "-" };
            dataGridView1.Rows.Add(rowInfo);
            if (AllPass)
            {
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[6].Value = "PASS";
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].DefaultCellStyle.BackColor = Color.SkyBlue;
            }
            else
            {
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[6].Value = "FAIL";
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
            }

            if(sigPathName == "main_mic")
            {
                for (int i = 0; i < gCountFrequencyPoint; i++) // find nearst frequency and get measured value
                {
                    nearest_freq = FindSweepPointIndex(XValues_Left, Convert.ToDouble(gFrequencyPoint[i]));
                    gMeasureValueSweepLeft[i] = MeasureValueSweepLeft[i] = YValues_Left[nearest_freq];
                    //gMeasureValueSweepRight[i] = MeasureValueSweepRight[i] = YValues_Right[nearest_freq];

                    dataGridView2.Rows[i].Cells[sweepResultPosition].Value = MeasureValueSweepLeft[i].ToString("F3");
                    //dataGridView2.Rows[i].Cells[3].Value = MeasureValueSweepRight[i].ToString("F3");
                }

            }
            else if (sigPathName == "sub_mic")
            {
                for (int i = 0; i < gCountFrequencyPoint; i++) // find nearst frequency and get measured value
                {
                    nearest_freq = FindSweepPointIndex(XValues_Left, Convert.ToDouble(gFrequencyPoint[i]));
                    gMeasureValueSweepSub[i] = MeasureValueSweepLeft[i] = YValues_Left[nearest_freq];
                    //gMeasureValueSweepRight[i] = MeasureValueSweepRight[i] = YValues_Right[nearest_freq];

                    dataGridView2.Rows[i].Cells[sweepResultPosition].Value = MeasureValueSweepLeft[i].ToString("F3");
                    //dataGridView2.Rows[i].Cells[3].Value = MeasureValueSweepRight[i].ToString("F3");
                }

            }
            if (AllPass) return true;
            else return false;
        }

        private bool checkLogSubItem(string measurementName)
        {
            /*
             * this temporary function will return
             * true : if measurement is not Sweep
             * false : if measurement is Sweep
             * for log function
             */
            bool x = false;
            switch (measurementName)
            {
                case "Level and Gain":
                    x = true;
                    break;
                case "THD+N":
                    x = true;
                    break;
                case "Signal to Noise Ratio":
                    x = true;
                    break;
                case "Stepped Frequency Sweep":
                    x = false;
                    break;
                case "Frequency Response":
                    x = false;
                    break;
                default:
                    break;
            }

            return x;
        }
        
#region CheckLogFile
        private bool checkLogFile()
        {
            DateTime today = DateTime.Now;
            string strDatePrefix = today.ToString("yy-MM-dd");
            string strLogFile = strDatePrefix + "_" + gModelName + "_PBA_Audio_Test" + "-log.csv";

            // set file path
            pathLogFile = System.IO.Directory.GetCurrentDirectory() + "\\" + "log\\" + strLogFile;

            if (File.Exists(pathLogFile))
            {
                // re-set test Cound
                StreamReader sr = new StreamReader(pathLogFile);
                string line;
                UInt32 lineCount = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    lineCount++;
                }
                sr.Close();

                return true;
            }
            else
            {
                //make new log file
                StreamWriter sw = new StreamWriter(pathLogFile, true, Encoding.Unicode);

                sw.Write("date\t result\t BD\t");  
                //sw.Write("result,");

                foreach (PathAndMeasurement seq in gMeasurementList)
                {
                    if (checkLogSubItem(seq.Meas.Name))
                    {
                        // this measurement items need left, right result
                        //sw.Write(seq.Meas.Name + " (Left),");
                        //sw.Write(seq.Meas.Name + " (Right),");
                        //sw.Write(seq.Meas.Name + " (Left)\t" + seq.Meas.Name + " (Right)\t");
                        sw.Write(seq.Meas.Name + " (" + seq.Path.Name + ")" + "\t");
                    }
                    else
                    {
                        // this measurement items (stepped frequency response) need frequency information also
                        //if(seq.Meas.Name == "Frequency Response")  // KS 2020-05-25 KIARA
                        if (seq.Meas.Name == "Stepped Frequency Sweep")
                        {
                            for (int i = 0; i < gCountFrequencyPoint; i++)
                            {
                                //sw.Write("Sweep_" + gFrequencyPoint[i] + " (Left),");
                                //sw.Write("Sweep_" + gFrequencyPoint[i] + " (Right),");
                                //sw.Write("Sweep_" + gFrequencyPoint[i] + " (Left)\t" + "Sweep_" + gFrequencyPoint[i] + " (Right) \t");
                                sw.Write("Sweep_" + gFrequencyPoint[i] + " (" + seq.Path.Name + ")" + "\t");
                            }
                        }
                        if (seq.Meas.Name == "Frequency Response")
                        {
                            for (int i = 0; i < gCountFrequencyPoint; i++)
                            {
                                //sw.Write("Sweep_" + gFrequencyPoint[i] + " (Left),");
                                //sw.Write("Sweep_" + gFrequencyPoint[i] + " (Right),");
                                //sw.Write("Sweep_" + gFrequencyPoint[i] + " (Left)\t" + "Sweep_" + gFrequencyPoint[i] + " (Right) \t");
                                sw.Write("Sweep_" + gFrequencyPoint[i] + " (" + seq.Path.Name + ")" + "\t");
                            }
                        }
                    }
                }

                sw.WriteLine("");
                sw.Close();

                return false;
            }
        }
#endregion CheckLogFile

#region WriteLog
        private void writeLog(bool ng_result) 
        {
            string date = DateTime.Now.ToString("yy-MM-dd HH:mm:ss");

            // open stream
            StreamWriter sw = new StreamWriter(pathLogFile, true, Encoding.Unicode);

            //sw.Write(date + ",");
            //sw.Write(ng_result ? "FAIL," : "PASS,");

            sw.Write(date + "\t");
            sw.Write(ng_result ? "FAIL \t" : "PASS \t");
            sw.Write(dutFullBdAddressAgent + "\t");

            foreach (PathAndMeasurement seq in gMeasurementList)
            {
                if (checkLogSubItem(seq.Meas.Name))
                {
                    if (seq.Path.Name == "main_mic")
                    {
                        if (seq.Meas.Name == "Level and Gain")
                        {
                            sw.Write(gMeasureValueLevelLeft.ToString("F3") + "\t");
                            //sw.Write(gMeasureValueLevelRight.ToString("F3") + "\t");
                        }
                        else if (seq.Meas.Name == "THD+N")
                        {
                            sw.Write(gMeasureValueThdLeft.ToString("F3") + "\t");
                            //sw.Write(gMeasureValueThdRight.ToString("F3") + "\t");
                        }
                        else if (seq.Meas.Name == "Signal to Noise Ratio")
                        {
                            sw.Write(gMeasureValueSnrLeft.ToString("F3") + "\t");
                            //sw.Write(gMeasureValueSnrRight.ToString("F3") + "\t");
                        }
                    }
                    else if (seq.Path.Name == "sub_mic")
                    {
                        if (seq.Meas.Name == "Level and Gain")
                        {
                            sw.Write(gMeasureValueLevelSub.ToString("F3") + "\t");
                            //sw.Write(gMeasureValueLevelRight.ToString("F3") + "\t");
                        }
                        else if (seq.Meas.Name == "THD+N")
                        {
                            sw.Write(gMeasureValueThdSub.ToString("F3") + "\t");
                            //sw.Write(gMeasureValueThdRight.ToString("F3") + "\t");
                        }
                        else if (seq.Meas.Name == "Signal to Noise Ratio")
                        {
                            sw.Write(gMeasureValueSnrSub.ToString("F3") + "\t");
                            //sw.Write(gMeasureValueSnrRight.ToString("F3") + "\t");
                        }
                    }
                }
                else
                {
                    // this measurement items (stepped frequency response) need frequency information also
                    if (seq.Path.Name == "main_mic")
                    {
                        if (seq.Meas.Name == "Stepped Frequency Sweep")
                        {
                            for (int i = 0; i < gCountFrequencyPoint; i++)
                            {
                                sw.Write(gMeasureValueSweepLeft[i].ToString("F3") + "\t");
                                //sw.Write(gMeasureValueSweepRight[i].ToString("F3") + "\t");
                            }
                        }

                        if (seq.Meas.Name == "Frequency Response")
                        {
                            for (int i = 0; i < gCountFrequencyPoint; i++)
                            {
                                sw.Write(gMeasureValueSweepLeft[i].ToString("F3") + "\t");
                                //sw.Write(gMeasureValueSweepRight[i].ToString("F3") + "\t");
                            }
                        }
                    }
                    else if (seq.Path.Name == "sub_mic")
                    {
                        if (seq.Meas.Name == "Stepped Frequency Sweep")
                        {
                            for (int i = 0; i < gCountFrequencyPoint; i++)
                            {
                                sw.Write(gMeasureValueSweepSub[i].ToString("F3") + "\t");
                                //sw.Write(gMeasureValueSweepRight[i].ToString("F3") + "\t");
                            }
                        }

                        if (seq.Meas.Name == "Frequency Response")
                        {
                            for (int i = 0; i < gCountFrequencyPoint; i++)
                            {
                                sw.Write(gMeasureValueSweepSub[i].ToString("F3") + "\t");
                                //sw.Write(gMeasureValueSweepRight[i].ToString("F3") + "\t");
                            }
                        }
                    }
                }
            }

            sw.WriteLine("");

            // close stream
            sw.Close();
        }

        #endregion WriteLog

        #region UART_JIG_CONTROL
        private void vbat_on()
        {
            int flag_ng = 0;

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                // Open first device in our list by serial number    
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialVbatComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBitMode(0xF1, 0x20);
                Thread.Sleep(200);

                ftStatus = myFtdiDevice.SetBitMode(0x00, 0x00);

                ftStatus = myFtdiDevice.Close();
            }
        }

        private void vbat_off()
        {
            int flag_ng = 0;

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialVbatComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBitMode(0xF0, 0x20);
                Thread.Sleep(200);

                ftStatus = myFtdiDevice.SetBitMode(0x00, 0x00);

                ftStatus = myFtdiDevice.Close();
            }
        }

        private void key_on()
        {
            int flag_ng = 0;

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialKeyComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBitMode(0xF1, 0x20);
                //Thread.Sleep(200);

                ftStatus = myFtdiDevice.SetBitMode(0x00, 0x00);

                ftStatus = myFtdiDevice.Close();
            }
        }

        private void key_off()
        {
            int flag_ng = 0;

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialKeyComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBitMode(0xF0, 0x20);
                //Thread.Sleep(200);
                ftStatus = myFtdiDevice.SetBitMode(0x00, 0x00);

                ftStatus = myFtdiDevice.Close();
            }
        }

        private static DateTime Delay(int MS)
        {
            DateTime ThisMoment = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime AfterWards = ThisMoment.Add(duration);
            while (AfterWards >= ThisMoment)
            {
                System.Windows.Forms.Application.DoEvents();
                ThisMoment = DateTime.Now;
            }

            return DateTime.Now;

        }
#endregion //#region UART_JIG_CONTROL

        private bool sendMainMicLoopbackCmd()
        {
            int timeout = 1500;
            bool bFoundResponse = false;

            int flag_ng = 0;

            int status = 0;

            UInt32 numBytesWritten = 0;

            // byte[] cmd_run_test = new byte[] { 0x05, 0x5A, 0x04, 0x00, 0x01, 0x77, 0xEE, 0x06, 0xF0 }; // TWIG2
            byte[] cmd_run_test = new byte[] { 0x05, 0x5A, 0x04, 0x00, 0x01, 0x00, 0x0B, 0x01 }; // ATH-CKS50TW2

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialKeyComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set up device data parameters
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBaudRate(3000000);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set data characteristics - Data bits, Stop bits, Parity
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetDataCharacteristics(FTDI.FT_DATA_BITS.FT_BITS_8, FTDI.FT_STOP_BITS.FT_STOP_BITS_1, FTDI.FT_PARITY.FT_PARITY_NONE);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set flow control - set RTS/CTS flow control
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetFlowControl(FTDI.FT_FLOW_CONTROL.FT_FLOW_NONE, 0, 0);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            timeout = 5000;
            do
            {
                if (flag_ng != 1)
                {
                    // do test
                    ftStatus = myFtdiDevice.Write(cmd_run_test, cmd_run_test.Count(), ref numBytesWritten);

                    if (ftStatus != FTDI.FT_STATUS.FT_OK)
                    {
                        flag_ng = 1;
                    }
                    Thread.Sleep(200);
                }

                if (flag_ng != 1)
                {
                    UInt32 numBytesAvailable = 0;
                    {
                        ftStatus = myFtdiDevice.GetRxBytesAvailable(ref numBytesAvailable);
                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                        }
                    }

                    if (flag_ng != 1 && numBytesAvailable != 0)
                    {
                        Console.WriteLine("numBytesAvailable = {0} \r\n", numBytesAvailable);

                        // Now that we have the amount of data we want available, read it
                        byte[] readData = new byte[numBytesAvailable];

                        string[] strData = new string[numBytesAvailable];

                        UInt32 numBytesRead = 0;
                        // Note that the Read method is overloaded, so can read string or byte array data
                        ftStatus = myFtdiDevice.Read(readData, numBytesAvailable, ref numBytesRead);

                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                            break;
                        }
                        else
                        {

                            // decoding (0x77 0xEE => 0x11, 0x77 0xEC => 0x13, 0x77 0x88 => 0x77) as AIROHA material
                            byte[] needProc = new byte[numBytesAvailable];
                            int i = 0;
                            int j = 0;
                            for (i = 0; i < numBytesAvailable; i++)
                            {
                                if (readData[i] == 0x77)
                                {
                                    if (readData[i + 1] == 0xEE)
                                    {
                                        needProc[j] = 0x11;
                                    }
                                    else if (readData[i + 1] == 0xEC)
                                    {
                                        needProc[j] = 0x13;
                                    }
                                    else if (readData[i + 1] == 0x88)
                                    {
                                        needProc[j] = 0x77;
                                    }
                                    i++;
                                }
                                else
                                {
                                    needProc[j] = readData[i];
                                }
                                Console.Write("{0:X2} ", needProc[j]);
                                j++;
                            }
                            Console.WriteLine("\r\n");

                            for (i = 0; i < j; i++)
                            {
                                /* ATH-CKS50TW2 CHECK */
                                if ((j - i) > 2) // check length : (j - current position) > 2 : to read header 0x05, 0x5B
                                {
                                    if (needProc[i] == 0x05 && needProc[i + 1] == 0x5B)
                                    {
                                        if (((j - i) - 4) >= ((needProc[i + 3] * 256) + (needProc[i + 2]))) // check length : (j - current position - 4) >= response length : length is payload size so need to minus 4 byte (header, length) to avoid meet getting wrong memory address
                                        {
                                            if (needProc[i + 4] == 0x01 && needProc[i + 5] == 0x00 && needProc[i + 6] == 0x0B)
                                            {
                                                bFoundResponse = true;
                                                status = needProc[i + 6];
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (bFoundResponse)
                            {
                                Console.WriteLine("RESPONSE FOUND!!!!!!\r\n");
                                Console.WriteLine("STATUS = {0}\r\n", status);
                            }
                        }
                    }
                }

                if (bFoundResponse) { break; }

                timeout -= 100;
                Thread.Sleep(10);
            } while (timeout > 0);

            if (timeout == 0)
            {
                flag_ng = 1;
            }

            Thread.Sleep(500);

            ftStatus = myFtdiDevice.Close();

            if (flag_ng != 1)
                return true;

            return false;
        }

        private bool sendSubMicLoopbackCmd()
        {
            int timeout = 1500;
            bool bFoundResponse = false;

            int flag_ng = 0;

            int status = 0;

            UInt32 numBytesWritten = 0;
        
            byte[] cmd_run_test = new byte[] { 0x05, 0x5A, 0x04, 0x00, 0x01, 0x00, 0x0B, 0x02 }; // ATH-CKS50TW2

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialKeyComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set up device data parameters
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBaudRate(3000000);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set data characteristics - Data bits, Stop bits, Parity
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetDataCharacteristics(FTDI.FT_DATA_BITS.FT_BITS_8, FTDI.FT_STOP_BITS.FT_STOP_BITS_1, FTDI.FT_PARITY.FT_PARITY_NONE);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set flow control - set RTS/CTS flow control
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetFlowControl(FTDI.FT_FLOW_CONTROL.FT_FLOW_NONE, 0, 0);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            timeout = 5000;
            do
            {
                if (flag_ng != 1)
                {
                    // do test
                    ftStatus = myFtdiDevice.Write(cmd_run_test, cmd_run_test.Count(), ref numBytesWritten);

                    if (ftStatus != FTDI.FT_STATUS.FT_OK)
                    {
                        flag_ng = 1;
                    }
                    Thread.Sleep(200);
                }

                if (flag_ng != 1)
                {
                    UInt32 numBytesAvailable = 0;
                    {
                        ftStatus = myFtdiDevice.GetRxBytesAvailable(ref numBytesAvailable);
                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                        }
                    }

                    if (flag_ng != 1 && numBytesAvailable != 0)
                    {
                        Console.WriteLine("numBytesAvailable = {0} \r\n", numBytesAvailable);

                        // Now that we have the amount of data we want available, read it
                        byte[] readData = new byte[numBytesAvailable];

                        string[] strData = new string[numBytesAvailable];

                        UInt32 numBytesRead = 0;
                        // Note that the Read method is overloaded, so can read string or byte array data
                        ftStatus = myFtdiDevice.Read(readData, numBytesAvailable, ref numBytesRead);

                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                            break;
                        }
                        else
                        {

                            // decoding (0x77 0xEE => 0x11, 0x77 0xEC => 0x13, 0x77 0x88 => 0x77) as AIROHA material
                            byte[] needProc = new byte[numBytesAvailable];
                            int i = 0;
                            int j = 0;
                            for (i = 0; i < numBytesAvailable; i++)
                            {
                                if (readData[i] == 0x77)
                                {
                                    if (readData[i + 1] == 0xEE)
                                    {
                                        needProc[j] = 0x11;
                                    }
                                    else if (readData[i + 1] == 0xEC)
                                    {
                                        needProc[j] = 0x13;
                                    }
                                    else if (readData[i + 1] == 0x88)
                                    {
                                        needProc[j] = 0x77;
                                    }
                                    i++;
                                }
                                else
                                {
                                    needProc[j] = readData[i];
                                }
                                Console.Write("{0:X2} ", needProc[j]);
                                j++;
                            }
                            Console.WriteLine("\r\n");

                            for (i = 0; i < j; i++)
                            {
                                /* ATH-CKS50TW2 CHECK */
                                if ((j - i) > 2) // check length : (j - current position) > 2 : to read header 0x05, 0x5B
                                {
                                    if (needProc[i] == 0x05 && needProc[i + 1] == 0x5B)
                                    {
                                        if (((j - i) - 4) >= ((needProc[i + 3] * 256) + (needProc[i + 2]))) // check length : (j - current position - 4) >= response length : length is payload size so need to minus 4 byte (header, length) to avoid meet getting wrong memory address
                                        {
                                            if (needProc[i + 4] == 0x01 && needProc[i + 5] == 0x00 && needProc[i + 6] == 0x0B)
                                            {
                                                bFoundResponse = true;
                                                status = needProc[i + 6];
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (bFoundResponse)
                            {
                                Console.WriteLine("RESPONSE FOUND!!!!!!\r\n");
                                Console.WriteLine("STATUS = {0}\r\n", status);
                            }
                        }
                    }
                }

                if (bFoundResponse) { break; }

                timeout -= 100;
                Thread.Sleep(10);
            } while (timeout > 0);

            if (timeout == 0)
            {
                flag_ng = 1;
            }

            Thread.Sleep(500);

            ftStatus = myFtdiDevice.Close();

            if (flag_ng != 1)
                return true;

            return false;
        }

#if (READ_BD_ADDR_BEFORE_TEST)
        private bool openUartForBDCmd()
        {
            dutFullBdAddressAgent = "";

            int timeout = 1500;
            bool bFoundResponse = false;

            int flag_ng = 0;

            int status = 0;

            UInt32 numBytesWritten = 0;

            byte[] cmd_log_off = new byte[] { 0x05, 0x5A, 0x06, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00 };
            byte[] cmd_run_test = new byte[] { 0x05, 0x5A, 0x06, 0x00, 0x00, 0x0A, 0x00, 0x36, 0x06, 0x00 };

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialKeyComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set up device data parameters
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBaudRate(3000000);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set data characteristics - Data bits, Stop bits, Parity
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetDataCharacteristics(FTDI.FT_DATA_BITS.FT_BITS_8, FTDI.FT_STOP_BITS.FT_STOP_BITS_1, FTDI.FT_PARITY.FT_PARITY_NONE);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set flow control - set RTS/CTS flow control
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetFlowControl(FTDI.FT_FLOW_CONTROL.FT_FLOW_NONE, 0, 0);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            if (flag_ng != 1)
            {
                // send log off cmd
                ftStatus = myFtdiDevice.Write(cmd_log_off, cmd_log_off.Count(), ref numBytesWritten);

                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
                Thread.Sleep(200);
            }

            timeout = 5000;
            do
            {
                if (flag_ng != 1)
                {
                    // do test
                    ftStatus = myFtdiDevice.Write(cmd_run_test, cmd_run_test.Count(), ref numBytesWritten);

                    if (ftStatus != FTDI.FT_STATUS.FT_OK)
                    {
                        flag_ng = 1;
                    }
                    Thread.Sleep(200);
                }

                if (flag_ng != 1)
                {
                    UInt32 numBytesAvailable = 0;
                    {
                        ftStatus = myFtdiDevice.GetRxBytesAvailable(ref numBytesAvailable);
                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                        }
                    }

                    if (flag_ng != 1 && numBytesAvailable != 0)
                    {
                        Console.WriteLine("numBytesAvailable = {0} \r\n", numBytesAvailable);

                        // Now that we have the amount of data we want available, read it
                        byte[] readData = new byte[numBytesAvailable];

                        string[] strData = new string[numBytesAvailable];

                        UInt32 numBytesRead = 0;
                        // Note that the Read method is overloaded, so can read string or byte array data
                        ftStatus = myFtdiDevice.Read(readData, numBytesAvailable, ref numBytesRead);

                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                            break;
                        }
                        else
                        {

#if true // processing 0x77 (0x77 0xEE => 0x11, 0x77 0xEC => 0x13, 0x77 0x88 => 0x77)
                            byte[] needProc = new byte[numBytesAvailable];
                            int i = 0;
                            int j = 0;
                            for (i = 0; i < numBytesAvailable; i++)
                            {
                                if (readData[i] == 0x77)
                                {
                                    if (readData[i + 1] == 0xEE)
                                    {
                                        needProc[j] = 0x11;
                                    }
                                    else if (readData[i + 1] == 0xEC)
                                    {
                                        needProc[j] = 0x13;
                                    }
                                    else if (readData[i + 1] == 0x88)
                                    {
                                        needProc[j] = 0x77;
                                    }
                                    i++;
                                }
                                else
                                {
                                    needProc[j] = readData[i];
                                }
                                Console.Write("{0:X2} ", needProc[j]);
                                j++;
                            }
                            Console.WriteLine("\r\n");

#endif

                            for (i = 0; i < j; i++)
                            {
                                //Console.Write("{0:X2} ", needProc[i + 4]);
                                if ((j - i) > 2) // check length : (j - current position) > 2 : to read header 0x05, 0x5B
                                {
                                    if (needProc[i] == 0x05 && needProc[i + 1] == 0x5B)
                                    {
                                        if (((j - i) - 4) >= ((needProc[i + 3] * 256) + (needProc[i + 2]))) // check length : (j - current position - 4) >= response length : length is payload size so need to minus 4 byte (header, length) to avoid meet getting wrong memory address
                                        {
                                            if (needProc[i + 5] == 0x0A && needProc[i + 6] == 0x06)
                                            {
                                                bFoundResponse = true;
                                                status = needProc[i + 7];
                                                //err_times = (needProc[i + 8] * 256) + (needProc[i + 7]);
                                                //min_ppm = ((int)(needProc[i + 9]) & 0x000000ff) | ((int)(needProc[i + 10] << 8)) | ((int)(needProc[i + 11] << 16)) | ((int)(needProc[i + 12] << 24));
                                                //max_ppm = ((int)(needProc[i + 13]) & 0x000000ff) | ((int)(needProc[i + 14] << 8)) | ((int)(needProc[i + 15] << 16)) | ((int)(needProc[i + 16] << 24));

                                                Console.WriteLine("_____[kimgh] got Agent BD adress : {0} {1} {2} {3} {4} {5} ",
                                                        (needProc[i + 13]).ToString("X2"), (needProc[i + 12]).ToString("X2"), (needProc[i + 11]).ToString("X2"), (needProc[i + 10]).ToString("X2"),
                                                        (needProc[i + 9]).ToString("X2"), (needProc[i + 8]).ToString("X2"));

                                                dutFullBdAddressAgent = (needProc[i + 13]).ToString("X2")+ (needProc[i + 12]).ToString("X2")+ (needProc[i + 11]).ToString("X2")+ (needProc[i + 10]).ToString("X2")+
                                                        (needProc[i + 9]).ToString("X2")+ (needProc[i + 8]).ToString("X2");

                            break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (bFoundResponse)
                            {
                                Console.WriteLine("RESPONSE FOUND!!!!!!\r\n");
                                Console.WriteLine("STATUS = {0}\r\n", status);

                            }
                        }
                    }
                }

                if (bFoundResponse) { break; }

                timeout -= 100;
                Thread.Sleep(10);
            } while (timeout > 0);

            if (timeout == 0)
            {
                flag_ng = 1;
            }

            Thread.Sleep(500);

            ftStatus = myFtdiDevice.Close();

            if (flag_ng != 1)
                return true;

            return false;
        }
#endif //#if (READ_BD_ADDR_BEFORE_TEST)

        

#region TEST_BUTTON
        private void button8_Click(object sender, EventArgs e)
        {
            int timeout = 1500;
            bool bFoundResponse = false;

            int flag_ng = 0;

            int status = 0;

            UInt32 numBytesWritten = 0;

            byte[] cmd_run_test = new byte[] { 0x05, 0x5A, 0x04, 0x00, 0x01, 0x77, 0xEE, 0x06, 0xF0 };

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialKeyComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set up device data parameters
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBaudRate(3000000);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set data characteristics - Data bits, Stop bits, Parity
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetDataCharacteristics(FTDI.FT_DATA_BITS.FT_BITS_8, FTDI.FT_STOP_BITS.FT_STOP_BITS_1, FTDI.FT_PARITY.FT_PARITY_NONE);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set flow control - set RTS/CTS flow control
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetFlowControl(FTDI.FT_FLOW_CONTROL.FT_FLOW_NONE, 0, 0);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            timeout = 5000;
            do
            {
                if (flag_ng != 1)
                {
                    // do test
                    ftStatus = myFtdiDevice.Write(cmd_run_test, cmd_run_test.Count(), ref numBytesWritten);

                    if (ftStatus != FTDI.FT_STATUS.FT_OK)
                    {
                        flag_ng = 1;
                    }
                    Thread.Sleep(200);
                }

                if (flag_ng != 1)
                {
                    UInt32 numBytesAvailable = 0;
                    {
                        ftStatus = myFtdiDevice.GetRxBytesAvailable(ref numBytesAvailable);
                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                        }
                    }

                    if (flag_ng != 1 && numBytesAvailable != 0)
                    {
                        Console.WriteLine("numBytesAvailable = {0} \r\n", numBytesAvailable);

                        // Now that we have the amount of data we want available, read it
                        byte[] readData = new byte[numBytesAvailable];

                        string[] strData = new string[numBytesAvailable];

                        UInt32 numBytesRead = 0;
                        // Note that the Read method is overloaded, so can read string or byte array data
                        ftStatus = myFtdiDevice.Read(readData, numBytesAvailable, ref numBytesRead);

                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                            break;
                        }
                        else
                        {
                            byte[] needProc = new byte[numBytesAvailable];
                            int i = 0;
                            int j = 0;
                            for (i = 0; i < numBytesAvailable; i++)
                            {
                                if (readData[i] == 0x77)
                                {
                                    if (readData[i + 1] == 0xEE)
                                    {
                                        needProc[j] = 0x11;
                                    }
                                    else if (readData[i + 1] == 0xEC)
                                    {
                                        needProc[j] = 0x13;
                                    }
                                    else if (readData[i + 1] == 0x88)
                                    {
                                        needProc[j] = 0x77;
                                    }
                                    i++;
                                }
                                else
                                {
                                    needProc[j] = readData[i];
                                }
                                Console.Write("{0:X2} ", needProc[j]);
                                j++;
                            }
                            Console.WriteLine("\r\n");

                            for (i = 0; i < j; i++)
                            {
                                //Console.Write("{0:X2} ", needProc[i + 4]);
                                if ((j - i) > 2) // check length : (j - current position) > 2 : to read header 0x05, 0x5B
                                {
                                    if (needProc[i] == 0x05 && needProc[i + 1] == 0x5B)
                                    {
                                        if (((j - i) - 4) >= ((needProc[i + 3] * 256) + (needProc[i + 2]))) // check length : (j - current position - 4) >= response length : length is payload size so need to minus 4 byte (header, length) to avoid meet getting wrong memory address
                                        {
                                            if (needProc[i + 4] == 0x01 && needProc[i + 5] == 0x11)
                                            {
                                                bFoundResponse = true;
                                                status = needProc[i + 6];
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (bFoundResponse)
                            {
                                Console.WriteLine("RESPONSE FOUND!!!!!!\r\n");
                                Console.WriteLine("STATUS = {0}\r\n", status);
                            }
                        }
                    }
                }

                if (bFoundResponse) { break; }

                timeout -= 100;
                Thread.Sleep(10);
            } while (timeout > 0);

            if (timeout == 0)
            {
                flag_ng = 1;
            }

            Thread.Sleep(500);

            ftStatus = myFtdiDevice.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int timeout = 1500;
            bool bFoundResponse = false;

            int flag_ng = 0;

            int status = 0;

            UInt32 numBytesWritten = 0;

            byte[] cmd_run_test = new byte[] { 0x05, 0x5A, 0x04, 0x00, 0x01, 0x77, 0xEE, 0x07, 0xF0 };

            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

            // Create new instance of the FTDI device class
            FTDI myFtdiDevice = new FTDI();

            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.OpenBySerialNumber(targetSerialKeyComport);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set up device data parameters
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetBaudRate(3000000);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set data characteristics - Data bits, Stop bits, Parity
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetDataCharacteristics(FTDI.FT_DATA_BITS.FT_BITS_8, FTDI.FT_STOP_BITS.FT_STOP_BITS_1, FTDI.FT_PARITY.FT_PARITY_NONE);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            // Set flow control - set RTS/CTS flow control
            if (flag_ng != 1)
            {
                ftStatus = myFtdiDevice.SetFlowControl(FTDI.FT_FLOW_CONTROL.FT_FLOW_NONE, 0, 0);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    flag_ng = 1;
                }
            }

            timeout = 5000;
            do
            {
                if (flag_ng != 1)
                {
                    // do test
                    ftStatus = myFtdiDevice.Write(cmd_run_test, cmd_run_test.Count(), ref numBytesWritten);

                    if (ftStatus != FTDI.FT_STATUS.FT_OK)
                    {
                        flag_ng = 1;
                    }
                    Thread.Sleep(200);
                }

                if (flag_ng != 1)
                {
                    UInt32 numBytesAvailable = 0;
                    {
                        ftStatus = myFtdiDevice.GetRxBytesAvailable(ref numBytesAvailable);
                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                        }
                    }

                    if (flag_ng != 1 && numBytesAvailable != 0)
                    {
                        Console.WriteLine("numBytesAvailable = {0} \r\n", numBytesAvailable);

                        // Now that we have the amount of data we want available, read it
                        byte[] readData = new byte[numBytesAvailable];

                        string[] strData = new string[numBytesAvailable];

                        UInt32 numBytesRead = 0;
                        // Note that the Read method is overloaded, so can read string or byte array data
                        ftStatus = myFtdiDevice.Read(readData, numBytesAvailable, ref numBytesRead);

                        if (ftStatus != FTDI.FT_STATUS.FT_OK)
                        {
                            flag_ng = 1;
                            break;
                        }
                        else
                        {
                            byte[] needProc = new byte[numBytesAvailable];
                            int i = 0;
                            int j = 0;
                            for (i = 0; i < numBytesAvailable; i++)
                            {
                                if (readData[i] == 0x77)
                                {
                                    if (readData[i + 1] == 0xEE)
                                    {
                                        needProc[j] = 0x11;
                                    }
                                    else if (readData[i + 1] == 0xEC)
                                    {
                                        needProc[j] = 0x13;
                                    }
                                    else if (readData[i + 1] == 0x88)
                                    {
                                        needProc[j] = 0x77;
                                    }
                                    i++;
                                }
                                else
                                {
                                    needProc[j] = readData[i];
                                }
                                Console.Write("{0:X2} ", needProc[j]);
                                j++;
                            }
                            Console.WriteLine("\r\n");

                            for (i = 0; i < j; i++)
                            {
                                //Console.Write("{0:X2} ", needProc[i + 4]);
                                if ((j - i) > 2) // check length : (j - current position) > 2 : to read header 0x05, 0x5B
                                {
                                    if (needProc[i] == 0x05 && needProc[i + 1] == 0x5B)
                                    {
                                        if (((j - i) - 4) >= ((needProc[i + 3] * 256) + (needProc[i + 2]))) // check length : (j - current position - 4) >= response length : length is payload size so need to minus 4 byte (header, length) to avoid meet getting wrong memory address
                                        {
                                            if (needProc[i + 4] == 0x01 && needProc[i + 5] == 0x11)
                                            {
                                                bFoundResponse = true;
                                                status = needProc[i + 6];
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (bFoundResponse)
                            {
                                Console.WriteLine("RESPONSE FOUND!!!!!!\r\n");
                                Console.WriteLine("STATUS = {0}\r\n", status);
                            }
                        }
                    }
                }

                if (bFoundResponse) { break; }

                timeout -= 100;
                Thread.Sleep(10);
            } while (timeout > 0);

            if (timeout == 0)
            {
                flag_ng = 1;
            }

            Thread.Sleep(500);

            ftStatus = myFtdiDevice.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            key_on();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            key_off();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            vbat_on();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            vbat_off();
        }
    }
}
#endregion // TEST_BUTTON


