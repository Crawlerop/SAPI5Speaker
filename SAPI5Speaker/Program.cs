using System;
using System.Text;
#if !USE_SPEECHLIB
using System.Speech.Synthesis;
#endif
using System.IO;
using System.Reflection;
#if USE_SPEECHLIB
using SpeechLib;
using System.Linq;
using System.Globalization;
#endif

namespace SAPI5Speaker
{
    internal class Program
    {
        static void Main(string[] args)
        {
            #if USE_SPEECHLIB
            SpVoice TTS = new SpVoice();
            #else
            SpeechSynthesizer TTS = new SpeechSynthesizer();
            #endif

            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;

            if (args.Length < 2)
            {
                if (args.Length == 1 && args[0] == "list")
                {
#if USE_SPEECHLIB
                    foreach (SpObjectToken voice in TTS.GetVoices())
                    {
                        Console.WriteLine(voice.GetAttribute("Name"));                        
                        Console.WriteLine("  Description: " + voice.GetDescription(0));
                        try
                        {
                            Console.WriteLine("  Age: " + voice.GetAttribute("Age"));
                        } catch
                        {
                            Console.WriteLine("  Age: unknown");
                        }
                        Console.WriteLine("  LCID: " + voice.GetAttribute("Language"));
                        try {
                            Console.WriteLine("  Language: " + CultureInfo.GetCultureInfo(int.Parse(voice.GetAttribute("Language").Split((char)0x3b)[0], System.Globalization.NumberStyles.HexNumber)));
                        } catch
                        {
                            Console.WriteLine("  Language: na-NA");
                        }
                        Console.WriteLine("  Gender: " + voice.GetAttribute("Gender"));
                        try
                        {
                            Console.WriteLine("  Vendor: " + voice.GetAttribute("Vendor"));
                        } catch
                        {
                            Console.WriteLine("  Vendor: N/A");
                        }
                        string[] regPath = voice.Id.Split(Path.DirectorySeparatorChar);
                        Console.WriteLine("  ID: " + (regPath[regPath.Length - 2] == "Tokens" ? regPath.Last() : String.Join("-", regPath.Skip(regPath.Length - 2).ToArray())));
                    }
#else
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
#endif
#if !USE_SPEECHLIB
                }
                else if (args.Length == 1 && args[0] == "test")
                {                    
                    foreach (InstalledVoice voice in TTS.GetInstalledVoices())
                    {                        
                        try
                        {
                            TTS.SelectVoice(voice.VoiceInfo.Name);
                            Console.WriteLine("{0}: OK", voice.VoiceInfo.Name);
                        } catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            Console.WriteLine("{0}: NOK", voice.VoiceInfo.Name);
                        }
                    }
#endif
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
                //Console.WriteLine(args[0]);
#if USE_SPEECHLIB
                SpObjectToken voiceToUse = null;
                foreach (SpObjectToken voice in TTS.GetVoices())
                {
                    if (voice.GetAttribute("Name") == args[0])
                    {
                        voiceToUse = voice;
                        break;
                    }
                }
                if (voiceToUse == null)
                {
                    Console.Error.WriteLine("That voice doesn't seems to be exist.");
                    Environment.Exit(1);
                }
                TTS.Voice = voiceToUse;
#else
                TTS.SelectVoice(args[0]);
#endif
                if (args[1] == "-")
                {
                    using (Stream stdout = Console.OpenStandardOutput())
                    {
                        using (MemoryStream buffer = new MemoryStream())
                        {
#if USE_SPEECHLIB
                            TTS.AudioOutputStream = (ISpeechMemoryStream)buffer;
#else
                            TTS.SetOutputToWaveStream(buffer);
#endif
                            String type = Console.ReadLine();
                            if (type == "s")
                            {
#if USE_SPEECHLIB
                                TTS.Speak(String.Format("<speak version=\"1.0\">{0}</speak>", Console.ReadLine()), SpeechVoiceSpeakFlags.SVSFParseSsml);
#else
                                TTS.SpeakSsml(String.Format("<speak version=\"1.0\">{0}</speak>", Console.ReadLine()));
#endif
                            }
                            else if (type == "t")
                            {
                                TTS.Speak(Console.ReadLine());
                            } else
                            {
                                Console.Error.WriteLine("Invalid type {0}", type);
                                Console.Error.WriteLine("Valid types include: t (text), s (SSML)");
                                Environment.Exit(1);
                            }                        

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
#if USE_SPEECHLIB
                    ISpeechFileStream outFile = (ISpeechFileStream)File.OpenWrite(args[1]);                    
                    TTS.AudioOutputStream = outFile;
#else
                    TTS.SetOutputToWaveFile(args[1]);
#endif

                    String type = Console.ReadLine();
                    if (type == "s")
                    {
#if USE_SPEECHLIB
                        TTS.Speak(String.Format("<speak version=\"1.0\">{0}</speak>", Console.ReadLine()), SpeechVoiceSpeakFlags.SVSFParseSsml);
#else
                        TTS.SpeakSsml(String.Format("<speak version=\"1.0\">{0}</speak>", Console.ReadLine()));
#endif
                    }
                    else if (type == "t")
                    {
                        TTS.Speak(Console.ReadLine());
                    }
                    else
                    {
                        Console.Error.WriteLine("Invalid type {0}", type);
                        Console.Error.WriteLine("Valid types include: t (text), s (SSML)");
                        Environment.Exit(1);
                    }
                }
            }
        }
    }
}
