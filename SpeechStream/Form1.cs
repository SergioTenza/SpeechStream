using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
//
using System.Speech.Recognition;
//
using NAudio.CoreAudioApi;
//
using System.Runtime.InteropServices;
//
using System.Threading;
//
using WindowsInput;
using WindowsInput.Native;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace SpeechStream
{
    public partial class StreamSpeechAssistan : Form
    {
        //
        // 
        //
        // import the function in your class
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);
        [DllImport("user32.dll")]
        static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);
        //
        //
        const int WM_CHAR = 0x105;
        //
        SpeechRecognitionEngine speechRecognitionEngine = null;
        List<String> words = new List<String>();
        string _culturePref = "es-ES";

        IntPtr _obsMain;        
        string _obsMain_Title;
        string _modifier = "MAY";
        int _obsMainId;

        bool canDetect = true;

        private System.Windows.Forms.Timer timer1;
        readonly int _time = 20;
        int _timeInterval = 1000;
        int _timeWorking = 0;
        //
        // 
        //
        public StreamSpeechAssistan()
        {
            InitializeComponent();
            buttonStart.Enabled = true;
            buttonStop.Enabled = false;
            textBoxMain.AppendText(Environment.NewLine + Environment.NewLine + "Welcome to SpeechStream Assistant. Just Talk and enjoy." + Environment.NewLine);            
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            

            var enumerator = new MMDeviceEnumerator();
            foreach (var endpoint in
                     enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
            {
                textBoxMain.AppendText( "Name: " + Convert.ToString(endpoint.FriendlyName) + Environment.NewLine + "ID: " + Convert.ToString(endpoint.ID) + Environment.NewLine);
                comboxMicros.Items.Add(Convert.ToString(endpoint.FriendlyName));
            }
            comboxMicros.SelectedIndex = 0;

            label7.Text = "Time until next action: " + _time.ToString();
            _timeWorking = _time;
        }
        //
        //
        //
        #region Start Speech Recognition Engine
        private SpeechRecognitionEngine createSpeechEngine(string preferredCulture)
        {
            foreach (RecognizerInfo config in SpeechRecognitionEngine.InstalledRecognizers())
            {
                if (config.Culture.ToString() == preferredCulture)
                {
                    textBoxMain.AppendText(SpeechRecognitionEngine.InstalledRecognizers()[0].Culture.ToString() + Environment.NewLine);
                    speechRecognitionEngine = new SpeechRecognitionEngine(config);
                    break;
                }
            }

            // if the desired culture is not found, then load default
            if (speechRecognitionEngine == null)
            {
                textBoxMain.AppendText("The desired culture is not installed on this machine, the speech-engine will continue using "+ Environment.NewLine
                    + SpeechRecognitionEngine.InstalledRecognizers()[0].Culture.ToString()+ " as the default culture."+Environment.NewLine+
                    "Culture " + preferredCulture + " not found!" + Environment.NewLine);
                speechRecognitionEngine = new SpeechRecognitionEngine(SpeechRecognitionEngine.InstalledRecognizers()[0]);
            }

            return speechRecognitionEngine;
        }
        private void loadGrammarAndCommands()
        {
            try
            {
                Choices texts = new Choices();
                string[] lines = File.ReadAllLines(Environment.CurrentDirectory + "\\example.txt");
                var count = 0;
                foreach (string line in lines)
                {
                    // skip commentblocks and empty lines..
                    if (line.StartsWith("--") || line == String.Empty) continue;

                    // split the line
                    var parts = line.Split(new char[] { '|' });

                    // add commandItem to the list for later lookup or execution
                    try
                    {
                        if (parts[0] == "voice")
                        {
                            textBoxMain.AppendText(String.Format("IsVoice= " + (parts[0] == "voice") + Environment.NewLine + 
                                                              "AttachedText = " + parts[1] + Environment.NewLine + 
                                                              "Modifier = " + parts[2] + Environment.NewLine +
                                                              "Assigned Key = " + parts[3] + Environment.NewLine +
                                                              "Return to = " + parts[4] + Environment.NewLine +
                                                              "Time in MiliSeconds= " + parts[5]) + Environment.NewLine);
                        }                  
                    }
                    catch (Exception ex)
                    {
                        textBoxMain.AppendText(ex + Environment.NewLine);
                    }


                    // add the text to the known choices of speechengine                    
                    texts.Add(parts[1]);
                    if (parts[0] == "voice")
                    {
                        words.Add(parts[0] + "|" + parts[1] + "|" + parts[2] + "|" + parts[3] + "|" + parts[4] + "|" + parts[5]);

                        if (parts[2] == "MAY")
                        {
                            _modifier = parts[2];
                        }
                        else if (parts[2] == "ALT")
                        {
                            _modifier = parts[2];
                        }
                        else if (parts[2] == "CTRL")
                        {
                            _modifier = parts[2];
                        }
                        switch (count)
                        {
                            case 0:
                                textBox1.Text = parts[1];
                                textBoxOn1.Text = parts[2];
                                textBoxMs1.Text = parts[5];
                                buttonAdd_1.Enabled = false;
                                count++;
                                break;
                            case 1:
                                textBox2.Text = parts[1];
                                textBoxOn2.Text = parts[2];
                                textBoxMs2.Text = parts[5];
                                count++;
                                break;
                            case 2:
                                textBox3.Text = parts[1];
                                textBoxOn3.Text = parts[2];
                                textBoxMs3.Text = parts[5];
                                count++;
                                break;
                            case 3:
                                textBox4.Text = parts[1];
                                textBoxOn4.Text = parts[2];
                                textBoxMs4.Text = parts[5];
                                count++;
                                break;
                            case 4:
                                textBox5.Text = parts[1];
                                textBoxOn5.Text = parts[2];
                                textBoxMs5.Text = parts[5];
                                count++;
                                break;
                            case 5:
                                textBox6.Text = parts[1];
                                textBoxOn6.Text = parts[2];
                                textBoxMs6.Text = parts[5];
                                count++;
                                break;
                            case 6:
                                textBox7.Text = parts[1];
                                textBoxOn7.Text = parts[2];
                                textBoxMs7.Text = parts[5];
                                count++;
                                break;
                            case 7:
                                textBox8.Text = parts[1];
                                textBoxOn8.Text = parts[2];
                                textBoxMs8.Text = parts[5];
                                count++;
                                break;
                            case 8:
                                textBox9.Text = parts[1];
                                textBoxOn9.Text = parts[2];
                                textBoxMs9.Text = parts[5];
                                count++;
                                break;
                            case 9:
                                textBox10.Text = parts[1];
                                textBoxOn10.Text = parts[2];
                                textBoxMs10.Text = parts[5];
                                count++;
                                break;
                            case 10:
                                textBox11.Text = parts[1];
                                textBoxOn11.Text = parts[2];
                                textBoxMs11.Text = parts[5];
                                count++;
                                break;
                            case 11:
                                textBox12.Text = parts[1];
                                textBoxOn12.Text = parts[2];
                                textBoxMs12.Text = parts[5];
                                count++;
                                break;
                            case 12:
                                textBox13.Text = parts[1];
                                textBoxOn13.Text = parts[2];
                                textBoxMs13.Text = parts[5];
                                count++;
                                break;
                            case 13:
                                textBox14.Text = parts[1];
                                textBoxOn14.Text = parts[2];
                                textBoxMs14.Text = parts[5];
                                count++;
                                break;
                            case 14:
                                textBox15.Text = parts[1];
                                textBoxOn15.Text = parts[2];
                                textBoxMs15.Text = parts[5];
                                count++;
                                break;
                            case 15:
                                textBox16.Text = parts[1];
                                textBoxOn16.Text = parts[2];
                                textBoxMs16.Text = parts[5];
                                count++;
                                break;
                            case 16:
                                textBox17.Text = parts[1];
                                textBoxOn17.Text = parts[2];
                                textBoxMs17.Text = parts[5];
                                count++;
                                break;
                            case 17:
                                textBox18.Text = parts[1];
                                textBoxOn18.Text = parts[2];
                                textBoxMs18.Text = parts[5];
                                count++;
                                break;
                            case 18:
                                textBox19.Text = parts[1];
                                textBoxOn19.Text = parts[2];
                                textBoxMs19.Text = parts[5];
                                count++;
                                break;
                            case 19:
                                textBox20.Text = parts[1];
                                textBoxOn20.Text = parts[2];
                                textBoxMs20.Text = parts[5];
                                count++;
                                break;
                            default:
                                break;
                        }
                    }
                }
                Grammar wordsList = new Grammar(new GrammarBuilder(texts));
                speechRecognitionEngine.LoadGrammar(wordsList);
                string[] _commands = words.ToArray();
                foreach (string line in _commands)
                { 
                    textBoxMain.AppendText(line + Environment.NewLine);
                }
                count = 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }        
        private void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //if (canDetect)
            //{
                //textBox1.AppendText("\r" + getKnownTextOrExecute(e.Result.Text) + Environment.NewLine);
            //}

            if (canDetect)
            {
                canDetect = false;                
                textBoxMain.AppendText("\r" + getKnownText(e.Result.Text) + Environment.NewLine);

                timer1 = new System.Windows.Forms.Timer();
                timer1.Tick += new EventHandler(timer1_Tick);
                timer1.Interval = _timeInterval; // 1 second
                timer1.Start();
                label7.Text = "Time until next action: " + _timeWorking.ToString();
            }
            
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            _timeWorking--;
            if (_timeWorking == 0)
            {
                timer1.Stop();
                canDetect = true;
                _timeWorking = _time;
            }                
            label7.Text = "Time until next action: " + _timeWorking.ToString();
        }

        private string getKnownTextOrExecute(string command)
        {
            try
            {
                foreach (string line in words)
                {
                    // split the line
                    var parts = line.Split(new char[] { '|' });

                    if (command == parts[0])
                    {
                        Process proc = new Process();
                        proc.EnableRaisingEvents = false;
                        proc.StartInfo.FileName = parts[1];
                        proc.Start();
                        return "You started: " + parts[0];
                    }                   
                }
                return "You said: " + command;
            }
            catch (Exception ex)
            {
                return "An error ocurred: " + Convert.ToString(ex);
            }
        }
        private string getKnownText(string command)
        {
            try
            {
                foreach (string line in words)
                {
                    // split the line
                    var parts = line.Split(new char[] { '|' });

                    if (command == parts[1])
                    {
                        textBoxMain.AppendText("Comando: " + parts[1] + Environment.NewLine);
                        if (parts[2] != "none")
                        {
                            textBoxMain.AppendText(" Enviando: " + parts[2] + " + " + parts[1] + Environment.NewLine);
                        }
                        else
                        {
                            textBoxMain.AppendText(" Enviando: " + parts[1] + Environment.NewLine);
                        }                            
                        ExecuteButton(parts[2], parts[3]);
                        Thread.Sleep(Convert.ToInt32(parts[5]));                        
                        textBoxMain.AppendText(" Enviando:  " +  parts[4] +  Environment.NewLine);
                        if (parts[4] == "main")
                        {
                            buttonMain.PerformClick();
                        }
                        else if (parts[4] == "afk")
                        {
                            buttonAFK.PerformClick();
                        }
                        else
                        {
                            textBoxMain.AppendText("Comando NO RECONOCIDO. REVISE TEXTO." + Environment.NewLine);
                        }
                        
                        return "You succesfully started: " + parts[1];                        
                    }
                }                
                return "You said: " + command;
            }
            catch (Exception ex)
            {                
                return "An error ocurred: " + Convert.ToString(ex);
            }
        }
        void engine_AudioLevelUpdated(object sender, AudioLevelUpdatedEventArgs e)
        {            
            progressBar1.Value = e.AudioLevel;
        }
        #endregion
        //
        //
        //
        #region General Buttons

        private void ButtonLaunchObs_Click(object sender, EventArgs e)
        {
            textBoxMain.AppendText("Lanzando OBS Studio..." + Environment.NewLine);
            ExecuteAsAdmin(@"c:\OBS Studio (64bit).lnk");
        }
        private void ButtonStart_Click(object sender, EventArgs e)
        {
            try
            {
                // create the engine
                speechRecognitionEngine = createSpeechEngine(_culturePref);

                // hook to events
                speechRecognitionEngine.AudioLevelUpdated += new EventHandler<AudioLevelUpdatedEventArgs>(engine_AudioLevelUpdated);
                speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(engine_SpeechRecognized);

                // load dictionary
                loadGrammarAndCommands();

                // use the system's default microphone
                speechRecognitionEngine.SetInputToDefaultAudioDevice();                

                // start listening
                speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Voice recognition failed");
            }
            buttonStart.Enabled = false;
            buttonStop.Enabled = true;
            textBoxMain.AppendText("Speech Engine Started. Now you can start Speaking..." + Environment.NewLine);            
        }

        private void ButtonStop_Click(object sender, EventArgs e)
        {
            if (speechRecognitionEngine != null)
            {
                // unhook events
                speechRecognitionEngine.RecognizeAsyncStop();
                // clean references
                speechRecognitionEngine.Dispose();
            }
            buttonStart.Enabled = true;
            buttonStop.Enabled = false;
            textBoxMain.AppendText("Speech Engine Stopped." + Environment.NewLine);          
        }

        private void ButtonClose_Click(object sender, EventArgs e)
        {
            string path = Environment.CurrentDirectory + "\\Handles.txt";
            File.WriteAllText(path, String.Empty);

            this.Close();
        }


        #endregion
        //
        // 
        //
        //        
        #region Idiomas
        private void ButtonEN_Click(object sender, EventArgs e)
        {
            if (buttonStart.Enabled)
            {
                _culturePref = "en-US";
                textBoxMain.AppendText("Your Selected Language is: " + _culturePref + Environment.NewLine);
            }
            else
            {
                MessageBox.Show("To change Language Please Push STOP Button First.");
            }
        }

        private void ButtonDE_Click(object sender, EventArgs e)
        {
            if (buttonStart.Enabled)
            {
                _culturePref = "de-DE";
                textBoxMain.AppendText("Your Selected Language is: " + _culturePref + Environment.NewLine);
            }
            else
            {
                MessageBox.Show("To change Language Please Push STOP Button First.");
            }
        }

        private void ButtonFR_Click(object sender, EventArgs e)
        {
            if (buttonStart.Enabled)
            {
                _culturePref = "fr-FR";
                textBoxMain.AppendText("Your Selected Language is: " + _culturePref + Environment.NewLine);
            }
            else
            {
                MessageBox.Show("To change Language Please Push STOP Button First.");
            }
        }

        private void ButtonJP_Click(object sender, EventArgs e)
        {
            if (buttonStart.Enabled)
            {
                _culturePref = "jp-JP";
                textBoxMain.AppendText("Your Selected Language is: " + _culturePref + Environment.NewLine);
            }
            else
            {
                MessageBox.Show("To change Language Please Push STOP Button First.");
            }
        }

        private void ButtonCN_Click(object sender, EventArgs e)
        {
            if (buttonStart.Enabled)
            {
                _culturePref = "zh-CN";
                textBoxMain.AppendText("Your Selected Language is: " + _culturePref + Environment.NewLine);
            }
            else
            {
                MessageBox.Show("To change Language Please Push STOP Button First.");
            }
        }

        private void ButtonTW_Click(object sender, EventArgs e)
        {
            if (buttonStart.Enabled)
            {
                _culturePref = "zh-TW";
                textBoxMain.AppendText("Your Selected Language is: " + _culturePref + Environment.NewLine);
            }
            else
            {
                MessageBox.Show("To change Language Please Push STOP Button First.");
            }
        }
        private void ButtonES_Click(object sender, EventArgs e)
        {
            if (buttonStart.Enabled)
            {
                _culturePref = "es-ES";
                textBoxMain.AppendText("Your Selected Language is: " + _culturePref + Environment.NewLine);
            }
            else
            {
                MessageBox.Show("To change Language Please Push STOP Button First.");
            }
        }
        #endregion
        //
        // 
        //
        #region Boton Add
        private void Button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    Choices texts = new Choices();
            //    string lines = textBoxFirst.Text + "|" + textBoxFirstDE.Text + '|' + textBoxFirstBool.Text;

            //        // split the line
            //        var parts = lines.Split(new char[] { '|' });

            //        // add commandItem to the list for later lookup or execution
            //        textBox1.AppendText(String.Format("Text = " + parts[0] + Environment.NewLine + "AttachedText = " + parts[1] + Environment.NewLine +
            //                                            "IsShellCommand = " + (parts[2] == "true") + Environment.NewLine));

            //        // add the text to the known choices of speechengine
            //        texts.Add(parts[0]);
            //        if (parts[2] == "true")
            //        {
            //            words.Add(parts[0] + "|" + parts[1]);
            //        }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Not Load Grammar Please check Formatting.");
            //}
        }        
        #endregion
        //
        // 
        //
        #region Botones Main    
        private void Button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_1);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_1);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_1);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_1);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_2);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_2);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_2);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_2);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_3);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_3);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_3);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_3);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_4);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_4);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_4);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_4);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_5);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_5);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_5);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_5);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }
        private void ButtonTryOff_1_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.F1);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.F1);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.F1);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.F1);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOff_2_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.F2);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.F2);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.F2);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.F2);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOff_3_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.F3);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.F3);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.F3);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.F3);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOff_4_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    IntPtr h = _obsMain;
                    SetForegroundWindow(h);
                    SendKeys.Send("{F9}");
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOff_5_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    IntPtr h = _obsMain;
                    SetForegroundWindow(h);
                    SendKeys.Send("{F12}");
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }
        private void ButtonTryOn_6_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_6);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_6);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_6);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_6);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOn_7_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_7);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_7);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_7);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_7);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOn_8_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_8);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_8);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_8);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_8);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOn_9_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_9);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_9);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_9);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_9);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonTryOn_10_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.VK_0);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_0);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_0);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_0);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }
        private void ButtonMain_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.F11);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.F11);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.F11);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.F11);
                    }                    
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }

        private void ButtonAFK_Click(object sender, EventArgs e)
        {
            try
            {
                if (_obsMain != null)
                {
                    if (_modifier == "MAY")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.F12);
                    }
                    else if (_modifier == "ALT")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.F12);
                    }
                    else if (_modifier == "CTRL")
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.F12);
                    }
                    else
                    {
                        IntPtr h = _obsMain;
                        SetForegroundWindow(h);
                        var sim = new InputSimulator();
                        sim.Keyboard.KeyDown(VirtualKeyCode.F12);
                    }
                }
                else
                {
                    textBoxMain.AppendText("Proceso Nulo..." + Environment.NewLine);
                }

            }
            catch (Exception Ex)
            {
                textBoxMain.AppendText(Ex + Environment.NewLine);
            }
        }
        #endregion
        //
        // 
        //
        private static void ShowDesktopWindows()
        {
            List<IntPtr> handles;
            List<string> titles;
            DesktopWindowsStuff.GetDesktopWindowHandlesAndTitles(out handles, out titles);
        }

        public void ExecuteAsAdmin(string fileName)
        {
            try
            {
                Process proc = new Process();
                proc.StartInfo.FileName = fileName;
                proc.Start();

                textBoxMain.AppendText("OBS Studio Lanzado..." + Environment.NewLine);

                Thread.Sleep(6000);                

                _obsMain = proc.MainWindowHandle;
                
                textBoxHandleMain.Text = Convert.ToString(proc.MainWindowHandle);
                textBoxMain.AppendText("Handle de Main asignado." + Environment.NewLine);

                _obsMain_Title = proc.MainWindowTitle;                
                textBoxMainTitle.Text = _obsMain_Title;
                _obsMainId = proc.Id;                

                Thread.Sleep(3000);                
            }
            catch (Exception e)
            {
                textBoxMain.AppendText(Convert.ToString(e));
            }
            
        }

        private void ExecuteButton(string button , string button2)
        {
            switch (button2)
            {
                case "1":
                    buttonTryOn_1.PerformClick();
                    break;
                case "2":
                    buttonTryOn_2.PerformClick();
                    break;
                case "3":
                    buttonTryOn_3.PerformClick();
                    break;
                case "4":
                    buttonTryOn_4.PerformClick();
                    break;
                case "5":
                    buttonTryOn_5.PerformClick();
                    break;
                case "6":
                    buttonTryOn_6.PerformClick();
                    break;
                case "7":
                    buttonTryOn_7.PerformClick();
                    break;
                case "8":
                    buttonTryOn_8.PerformClick();
                    break;
                case "9":
                    buttonTryOn_9.PerformClick();
                    break;
                case "0":
                    buttonTryOn_10.PerformClick();
                    break;
                case "F1":
                    buttonTryOn_11.PerformClick();
                    break;
                case "F2":
                    buttonTryOn_12.PerformClick();
                    break;
                case "F3":
                    buttonTryOn_13.PerformClick();
                    break;
                case "F4":
                    buttonTryOn_14.PerformClick();
                    break;
                case "F5":
                    buttonTryOn_15.PerformClick();
                    break;
                case "F6":
                    buttonTryOn_16.PerformClick();
                    break;
                case "F7":
                    buttonTryOn_17.PerformClick();
                    break;
                case "F8":
                    buttonTryOn_18.PerformClick();
                    break;
                case "F9":
                    buttonTryOn_19.PerformClick();
                    break;
                case "F10":
                    buttonTryOn_20.PerformClick();
                    break;
                default:
                    textBoxMain.AppendText("Comando no reconocido: " + button);
                    break;
            }
        }
        
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (speechRecognitionEngine != null && buttonStop.Enabled == true)
            {
                // unhook events
                speechRecognitionEngine.RecognizeAsyncStop();
                // clean references
                speechRecognitionEngine.Dispose();
            }
            string path = Environment.CurrentDirectory + "\\Handles.txt";
            File.WriteAllText(path, String.Empty);

            //if (_obsMain != null)
            //{
            //    try
            //    {
            //        Process proc = Process.GetProcessById(_obsMainId);
            //        proc.Kill();
            //    }
            //    catch (Exception)
            //    {

            //        throw;
            //    }                
            //}
        }


    }
}
