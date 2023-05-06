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
using System.Runtime.InteropServices.ComTypes;
#endif

namespace SAPI5Speaker
{
    internal class Program
    {
        static void Main(string[] args)
        {
            #if USE_SPEECHLIB
            SpVoice TTS = new SpVoice();
            TTS.AllowAudioOutputFormatChangesOnNextSet = true;
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
                        if (!voice.Enabled) continue;
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
#if FALSE
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
#if !USE_SPEECHLIB                       
                        using (MemoryStream buffer = new MemoryStream())
#endif
                        {
#if USE_SPEECHLIB
                            SpMemoryStream buffer = new SpMemoryStream();                            
                            TTS.AudioOutputStream = buffer;

                            String format = Console.ReadLine();
                            switch (format)
                            {
                                case "8":
                                    buffer.Format.Type = SpeechAudioFormatType.SAFT8kHz16BitMono;
                                    break;
                                case "11":
                                    buffer.Format.Type = SpeechAudioFormatType.SAFT11kHz16BitMono;
                                    break;
                                case "16":
                                    buffer.Format.Type = SpeechAudioFormatType.SAFT16kHz16BitMono;
                                    break;
                                case "22":
                                    buffer.Format.Type = SpeechAudioFormatType.SAFT22kHz16BitMono;
                                    break;
                                case "32":
                                    buffer.Format.Type = SpeechAudioFormatType.SAFT32kHz16BitMono;
                                    break;
                                case "44":
                                    buffer.Format.Type = SpeechAudioFormatType.SAFT44kHz16BitMono;
                                    break;
                                case "48":
                                    buffer.Format.Type = SpeechAudioFormatType.SAFT48kHz16BitMono;
                                    break;
                            }
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

#if USE_SPEECHLIB
                            buffer.Seek(0);
#else
                            buffer.Position = 0;
#endif

#if USE_SPEECHLIB
                            object temp = new object();
                            int count;
                            SpWaveFormatEx waveFmt = buffer.Format.GetWaveFormatEx();

                            //Console.Error.WriteLine(waveFmt.);
                            stdout.Write(Encoding.ASCII.GetBytes("RIFF"), 0, 4);
                            stdout.Write(BitConverter.GetBytes(((byte[])buffer.GetData()).Length + 0x24), 0, 4);
                            stdout.Write(Encoding.ASCII.GetBytes("WAVE"), 0, 4);

                            stdout.Write(Encoding.ASCII.GetBytes("fmt "), 0, 4);                            
                            stdout.Write(BitConverter.GetBytes(16), 0, 4);                            
                            stdout.Write(BitConverter.GetBytes(waveFmt.FormatTag), 0, 2);                            
                            stdout.Write(BitConverter.GetBytes(waveFmt.Channels), 0, 2);                            
                            stdout.Write(BitConverter.GetBytes(waveFmt.SamplesPerSec), 0, 4);
                            stdout.Write(BitConverter.GetBytes(waveFmt.AvgBytesPerSec), 0, 4);                            
                            stdout.Write(BitConverter.GetBytes(waveFmt.BlockAlign), 0, 2);                            
                            stdout.Write(BitConverter.GetBytes(waveFmt.BitsPerSample), 0, 2);

                            stdout.Write(Encoding.ASCII.GetBytes("data"), 0, 4);
                            stdout.Write(BitConverter.GetBytes(((byte[])buffer.GetData()).Length), 0, 4);

                            while ((count = buffer.Read(out temp, 8192)) > 0)
                            {
                                stdout.Write((byte[])temp, 0, count);
                            }

                            stdout.Flush();
#else
                            byte[] temp = new byte[8192];
                            int count;

                            while ((count = buffer.Read(temp, 0, (int)temp.Length)) > 0)
                            {
                                stdout.Write(temp, 0, count);
                            }

                            stdout.Flush();
#endif
                        }
                    }
                }
                else
                {
#if USE_SPEECHLIB
                    SpFileStream outFile = new SpFileStream();
                    outFile.Open(args[1], SpeechStreamFileMode.SSFMCreateForWrite, false);
                    TTS.AudioOutputStream = outFile;

                    String format = Console.ReadLine();
                    switch (format)
                    {
                        case "8":
                            outFile.Format.Type = SpeechAudioFormatType.SAFT8kHz16BitMono;
                            break;
                        case "11":
                            outFile.Format.Type = SpeechAudioFormatType.SAFT11kHz16BitMono;
                            break;
                        case "16":
                            outFile.Format.Type = SpeechAudioFormatType.SAFT16kHz16BitMono;
                            break;
                        case "22":
                            outFile.Format.Type = SpeechAudioFormatType.SAFT22kHz16BitMono;
                            break;                        
                        case "32":
                            outFile.Format.Type = SpeechAudioFormatType.SAFT32kHz16BitMono;
                            break;
                        case "44":
                            outFile.Format.Type = SpeechAudioFormatType.SAFT44kHz16BitMono;
                            break;
                        case "48":
                            outFile.Format.Type = SpeechAudioFormatType.SAFT48kHz16BitMono;
                            break;                        
                    }
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
