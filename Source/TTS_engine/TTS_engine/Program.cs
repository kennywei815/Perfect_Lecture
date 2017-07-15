using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Speech.Synthesis;
using System.Speech.AudioFormat;
using System.Globalization;


namespace TTS_engine
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                helpMsg();
                return -1;
            }

            string text_path = args[0];
            string wav_path = args[1];

            // Get current system language

            CultureInfo ci;
            ci = CultureInfo.InstalledUICulture;
            ci = CultureInfo.CurrentUICulture;
            ci = CultureInfo.CurrentCulture;

            string current_lang = ci.Name;

            // Initialize a new instance of the SpeechSynthesizer.
            SpeechSynthesizer synth = new SpeechSynthesizer();

            // Configure the audio output.
            synth.SetOutputToWaveFile(wav_path);

            // Speak a string.
            try
            {
                using (StreamReader text_file = new StreamReader(text_path))
                {
                    //[DONE]: set default value of xml:lang to system langauge

                    //string text = text_file.ReadToEnd();
                    string text = "<!-- ?xml version=\"1.0\"? --> \n"
                                + "<speak xmlns=\"http://www.w3.org/2001/10/synthesis\" \n"
                                + "       xmlns:dc=\"http://purl.org/dc/elements/1.1/\" \n"
                                + "       xml:lang=\"" + current_lang + "\" \n"
                                + "       version=\"1.0\"> \n"
                                + text_file.ReadToEnd()
                                + "\n"
                                + "</speak>";

                    //[TODO]: check if it is valid SSML script.
                    //[TODO]: process english words with xml:lang="en"

                    /*
                    MessageBox.Show(text, "執行語音合成指令...", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    */
                    synth.SpeakSsml(text);
                }
                
            }
            catch (Exception exception) { MessageBox.Show(exception.Message, "執行語音合成時發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            //Console.WriteLine();
            //Console.WriteLine("Press any key to exit...");
            //Console.ReadKey();

            return 0;
        }

        static void helpMsg()
        {
            const string help_text = "Usage:  TTS_engine.exe [text] [wav]";
            const string caption = "TTS_engine";
            var result = MessageBox.Show(help_text, caption);
        }
    }
}
