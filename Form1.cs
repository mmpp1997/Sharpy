using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Speech.Recognition;
using System.Speech.Synthesis;

namespace Sharpy
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine Spec = new SpeechRecognitionEngine();

        SpeechSynthesizer Sharpy = new SpeechSynthesizer();

        DateTime danas = DateTime.UtcNow.Date;
        public Form1()
        {
            Spec.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Start);
            Naredbe();
            Spec.SetInputToDefaultAudioDevice();
            Spec.RecognizeAsync(RecognizeMode.Multiple);
            InitializeComponent();
        }
      
        private void Naredbe()
        {
            Choices mogucnosti = new Choices();
            string[] naredbe = File.ReadAllLines(Environment.CurrentDirectory + "\\naredbe.txt");
            mogucnosti.Add(naredbe);
            Grammar rijecnik = new Grammar(new GrammarBuilder(mogucnosti));
            Spec.LoadGrammar(rijecnik);
        }
        private void Start(object sender, SpeechRecognizedEventArgs e)
        {
            string govor = e.Result.Text;          
            if (govor == "wake up" )
            {             
                Spec.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Spec_SpeechRecognized);
                Spec.SpeechRecognized -= new EventHandler<SpeechRecognizedEventArgs>(Start);

                Sharpy.Speak("how can i assist you today");
                textBox1.Text = "How can i Assist you today?";
                textBox2.Text = "Hint: Ask Sharpy: What can you do?";
                label1.Text = "Sharpy is now active";
                label1.BackColor = System.Drawing.Color.Green;
            }
        }
        private void Spec_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string govor = e.Result.Text;
            switch (govor)
            {
                case "whats up":
                    Sharpy.Speak("not much");
                    textBox1.Text = "What's up?";
                    textBox2.Text = "Not Much.";
                    break;

                case "can you open notepad":
                    Sharpy.Speak("sure just a second");
                    textBox1.Text = "Can you open notepad?";
                    textBox2.Text = "Sure, just a second.";
                    Process.Start("notepad.exe");
                    break;

                case "can you close notepad":
                    Sharpy.Speak("sure just a second");
                    textBox1.Text = "Can you close notepad?";
                    textBox2.Text = "Sure, just a second.";
                    Izvrsi("c:/nircmd.exe killprocess notepad.exe");
                    break;

                case "hide":
                    Sharpy.Speak("will do");
                    textBox1.Text = "Hide!";
                    textBox2.Text = "Will do!";
                    this.WindowState = FormWindowState.Minimized;
                    break;

                case "can you tell me the date":
                    Sharpy.Speak("sure today is " + danas.Date.ToString());
                    textBox1.Text = "Can you tell me the date?";
                    textBox2.Text = "Sure, Today is " + danas.Date.ToString();
                    break;

                case "how are you sharpy":
                    Sharpy.Speak("i am fine How are you");
                    textBox1.Text = "How are you Sharpy?";
                    textBox2.Text = "I'm fine. How are you?";
                    break;

                case "i am ok":
                    Sharpy.Speak("that is nice");
                    textBox1.Text = "Im OK!";
                    textBox2.Text = "That is nice";
                    break;

                

                
                    
                

                case "can you go to the web":
                    Sharpy.Speak("sure just a second");
                    textBox1.Text = "Can you go to the web?";
                    textBox2.Text = "Sure, just a second.";
                    Process.Start("MicrosoftEdge.exe");
                    break;

                case "can you close edge":
                    Sharpy.Speak("sure just a second");
                    textBox1.Text = "Can you close edge?";
                    textBox2.Text = "Sure, just a second.";
                    Izvrsi("c:/nircmd.exe killprocess MicrosoftEdge.exe");
                    break;

                case "change sound":
                    Sharpy.Speak("sure just a second");
                    textBox1.Text = "Change sound!";
                    textBox2.Text = "Sure, just a second.";
                    Izvrsi("C:/nircmd.exe mutesysvolume 2 ");
                    break;

                case "what did i copy":
                    Sharpy.Speak("You copied this");
                    textBox1.Text = "What did i copy?";
                    textBox2.Text = Clipboard.GetText(TextDataFormat.Text);
                    break;

                case "what can you do":
                    Sharpy.Speak("let me show you");
                    textBox1.Text = "What can you do?";
                    textBox2.Text = "Let me show you!";
                    Process.Start("notepad.exe", Environment.CurrentDirectory + "\\mogucnosti.txt");
                    break;

                case "that is all":
                    Sharpy.Speak("bye hear from you soon");
                    textBox1.Text = "That is all";
                    textBox2.Text = "Bye, hear from you soon!";
                    Application.Exit();
                    break;

                case "can you check the weather":
                    Sharpy.Speak("checking now");
                    textBox1.Text = "Can you check the weather?";
                    textBox2.Text = "Checking now!";
                    Process.Start("chrome.exe", "https://meteo.hr/prognoze.php?Code=Split&id=prognoza&section=prognoze_model&param=3d");
                    break;

                case "please open task manager":
                    Sharpy.Speak("i will");
                    textBox1.Text = "Please open task manager!";
                    textBox2.Text = "I will!";
                    Process.Start("Taskmgr.exe");
                    break;

                case "open quick access":
                    Sharpy.Speak("here you go");
                    textBox1.Text = "Open Quick access!";
                    textBox2.Text = "Here you go!";
                    Process.Start(@"C:\Users\Mario");
                    break;

                case "start":
                    Sharpy.Speak("on screen right now");
                    textBox1.Text = "Start!";
                    textBox2.Text = "On screen right now!";
                    SendKeys.Send("^{ESC}");                  
                    break;
     
                

                case "maximize":
                    Sharpy.Speak("will do");
                    textBox1.Text = "Maksimize!";
                    textBox2.Text = "Will do!";
                    this.WindowState = FormWindowState.Maximized;
                    break;

                case "open cmd":
                    Sharpy.Speak("opening cmd");
                    textBox1.Text = "open CMD!";
                    textBox2.Text = "Opening CMD";
                    Process.Start("cmd.exe");
                    break;
            }
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        public static void Izvrsi(string naredba)
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd", "/c" + naredba);
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.UseShellExecute = false;
            processStartInfo.CreateNoWindow = true;
            Process pro = new Process();
            pro.StartInfo = processStartInfo;

            pro.Start();
            pro.Close();
        }
    }
}
