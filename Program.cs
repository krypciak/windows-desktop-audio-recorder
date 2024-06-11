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
        }

        static void Main(string[] args)
        {
            var parser = new CommandLine.Parser(with => with.HelpWriter = null);
            var parserResult = parser.ParseArguments<Options>(args);
            parserResult
              .WithParsed<Options>(options => Run(options))
              .WithNotParsed(errs => DisplayHelp(parserResult, errs));
        }
        static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<CommandLine.Error> errs)
        {
            var helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Copyright = "";
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
                capture = null;
                audioWriter.Flush();
                audioStream.Flush();
                audioStream.Position = 0;

                using (var reader = new WaveFileReader(audioStream))
                {
                    MediaFoundationEncoder.EncodeToMp3(reader, path);
                }

                audioWriter.Dispose();
                audioStream.Dispose();
                wasapiOut.Dispose();
            };

            AppDomain.CurrentDomain.ProcessExit += new EventHandler((sender, e) =>
            {
                Console.WriteLine("Exiting...");
                capture.StopRecording();
            });

            capture.StartRecording();

            Console.WriteLine("Started recording");

            while (capture.CaptureState != NAudio.CoreAudioApi.CaptureState.Stopped)
            {
                Thread.Sleep(500);
            }
        }
    }
}
