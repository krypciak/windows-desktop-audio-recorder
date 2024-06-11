﻿using System.Diagnostics.CodeAnalysis;
using CommandLine;
using CommandLine.Text;
using NAudio.CoreAudioApi;
using NAudio.MediaFoundation;
using NAudio.Wave;

namespace DesktopAudioRecorder
{
    class Program
    {
        public class Options
        {
            [Option('o', "output", Required = true, HelpText = "Output audio file path")]
            public required string Output { get; set; }

            [Option('t', "time", Required = false, HelpText = "Recording length in seconds")]
            public int? Time { get; set; }
        }

        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(Options))]
        static void Main(string[] args)
        {
            var parser = new CommandLine.Parser(with => with.HelpWriter = null);
            var parserResult = parser.ParseArguments<Options>(args);
            parserResult
              .WithParsed<Options>(options => Run(options))
              .WithNotParsed(errs => DisplayHelp(parserResult, errs));
        }
        static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<Error> errs)
        {
            HelpText helpText = HelpText.AutoBuild(result, h =>
                {
                    h.AdditionalNewLineAfterOption = false;
                    h.AutoVersion = false;
                    return HelpText.DefaultParsingErrorsHandler(result, h);
                }, e => e);
            Console.WriteLine(helpText);
        }

        static private WasapiOut PlaySilence(WaveFormat format)
        {
            /* Play silence to fix WasapiLoopbackCapture not capturing silence */
            var device = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            var silenceProvider = new SilenceProvider(format);
            var wasapiOut = new WasapiOut(device, AudioClientShareMode.Shared, false, 250);
            wasapiOut.Init(silenceProvider);
            wasapiOut.Play();
            return wasapiOut;
        }

        static void Run(Options o)
        {
            var path = Path.GetFullPath(o.Output);
            var outputFolder = Path.GetDirectoryName(path)!;
            Directory.CreateDirectory(outputFolder);

            var capture = new WasapiLoopbackCapture();
            var audioStream = new MemoryStream(1024);
            var audioWriter = new WaveFileWriter(audioStream, capture.WaveFormat);

            WasapiOut wasapiOut = PlaySilence(capture.WaveFormat);

            capture.DataAvailable += (s, a) => audioWriter.Write(a.Buffer, 0, a.BytesRecorded);

            MediaFoundationApi.Startup();
            capture.RecordingStopped += (s, a) =>
            {
                capture.Dispose();
                audioWriter.Flush();
                audioStream.Flush();
                audioStream.Position = 0;

                Console.WriteLine("Encoding to mp3...");
                using (var reader = new WaveFileReader(audioStream))
                {
                    MediaFoundationEncoder.EncodeToMp3(reader, path);
                }
                Console.WriteLine("Output file:");
                Console.WriteLine(path);

                audioWriter.Dispose();
                audioStream.Dispose();
                wasapiOut.Dispose();
            };

            AppDomain.CurrentDomain.ProcessExit += new EventHandler((sender, e) =>
            {
                capture.StopRecording();
            });

            capture.StartRecording();

            Console.WriteLine("Started recording");

            DateTime? targetDate = o.Time != null ? DateTime.Now.AddSeconds((double)o.Time) : null;
            while (capture.CaptureState != NAudio.CoreAudioApi.CaptureState.Stopped)
            {
                Thread.Sleep(100);
                if (targetDate != null && DateTime.Now >= targetDate) { capture.StopRecording(); }
            }
        }
    }
}
