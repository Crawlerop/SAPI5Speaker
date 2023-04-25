using System;
using System.Collections.Generic;
using System.Text;
using System.Speech.Synthesis;
using System.IO;
using System.Reflection;

namespace SAPI5Speaker
{
    internal class Program
    {
        static void Main(string[] args)
        {
            SpeechSynthesizer TTS = new SpeechSynthesizer();

            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;

            if (args.Length <= 2)
            {
                if (args.Length == 1 && args[0] == "list")
                {                    
                    foreach (InstalledVoice voice in TTS.GetInstalledVoices())
                    {
                        Console.WriteLine(voice.VoiceInfo.Name);
                        Console.WriteLine("  Description: "+voice.VoiceInfo.Description);                        
                        Console.WriteLine("  Age: " + voice.VoiceInfo.Age);
                        Console.WriteLine("  Culture: " + voice.VoiceInfo.Culture);
                        Console.WriteLine("  Gender: " + voice.VoiceInfo.Gender);
                        Console.WriteLine("  ID: " + voice.VoiceInfo.Id);

                        Console.WriteLine("  AdditionalInfo: ");
                        
                        foreach (string key in voice.VoiceInfo.AdditionalInfo.Keys)
                        {
                            if (key.Length <= 0) continue;
                            Console.WriteLine(String.Format("    {0}: {1}", key, voice.VoiceInfo.AdditionalInfo[key]));
                        }                        
                    }
                }
                else
                {
                    Console.WriteLine("SAPI52WAV v" + Assembly.GetExecutingAssembly().GetName().Version.ToString());
                    Console.WriteLine("Usage: [0] Voice Output");
                    Console.WriteLine("\"[0] list\" to see the voices installed on this system.");
                }
            }
            else
            {
                TTS.SelectVoice(args[0]);
                if (args[1] == "-")
                {
                    using (Stream stdout = Console.OpenStandardOutput())
                    {
                        using (MemoryStream buffer = new MemoryStream())
                        {
                            TTS.SetOutputToWaveStream(buffer);
                            TTS.Speak(Console.ReadLine());

                            buffer.Position = 0;

                            byte[] temp = new byte[8192];
                            int count;

                            while ((count = buffer.Read(temp, 0, (int)temp.Length)) > 0)
                            {
                                stdout.Write(temp, 0, count);
                            }
                        }
                    }
                }
                else
                {
                    TTS.SetOutputToWaveFile(args[1]);
                    TTS.Speak(Console.ReadLine());
                }
            }
        }
    }
}
