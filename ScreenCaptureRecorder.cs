using System;
using System.Drawing;
using System.Threading.Tasks;
using Xabe.FFmpeg;
using Xabe.FFmpeg.Streams;

namespace WpfApp32
{
    internal class ScreenCaptureRecorder
    {
        public string OutputPath { get; internal set; }
        public VideoCodec VideoCodec { get; internal set; }
        public Rectangle CaptureRectangle { get; internal set; }

        public async Task StartRecordingAsync()
        {
            IVideoStream videoStream = new Xabe.FFmpeg.Streams.VideoStream(OutputPath, VideoCodec.h264);

            await FFmpeg.Conversions.New()
                .AddInput($"-f gdigrab -framerate 30 -i desktop")
                .AddStream(videoStream)
                .Start();
        }

        public async Task StopRecordingAsync()
        {
            await Task.Delay(1000); // Delay 1 second to ensure the recording is stopped.
            await FFmpeg.Conversions.New().StopRecordingAsync();
        }

        internal void Dispose()
        {
            throw new NotImplementedException();
        }

        internal void Start()
        {
            throw new NotImplementedException();
        }

        internal void Stop()
        {
            throw new NotImplementedException();
        }
    }
}
