using System;
using hyperionnet;
using FlatBuffers;
using static Humanizer.In;

namespace HyperionScreenCap.Networking
{
    class FbsClinet: HyperionClient
    {
        private bool prioritySet;

        public FbsClinet(string host, int port, int priority, int messageDuration) : base(host, port, priority, messageDuration)
        {
            prioritySet = false;
        }

        protected override void SendImageDataMessage(byte[] pixeldata, int width, int height)
        {
            if (!prioritySet)
            {
                SendPriorityRegistrationMessage();
                SendPriorityRegistrationMessage(); // Sending twice just in case a message errors out. TODO: de-dupe
                prioritySet = true;
            }
            var builder = new FlatBufferBuilder(1024);
            var rawImageDataOffset = RawImage.CreateDataVector(builder, pixeldata);
            var rawImageOffset = RawImage.CreateRawImage(builder, rawImageDataOffset, width, height);
            var imageOffset = Image.CreateImage(builder, ImageType.RawImage, rawImageOffset.Value, _messageDuration);
            var requestOffset = Request.CreateRequest(builder, Command.Image, imageOffset.Value);
            builder.Finish(requestOffset.Value);
            SendFinishedMessage(builder);
        }

        protected override void SendClearPriorityMessage()
        {
            var builder = new FlatBufferBuilder(64);
            var clearOffset = Clear.CreateClear(builder, _priority);
            var requestOffset = Request.CreateRequest(builder, Command.Clear, clearOffset.Value);
            builder.Finish(requestOffset.Value);
            SendFinishedMessage(builder);
        }

        private void SendPriorityRegistrationMessage()
        {
            var builder = new FlatBufferBuilder(64);
            var originOffset = builder.CreateString("HyperionScreenCap");
            var registerOffset = Register.CreateRegister(builder, originOffset, _priority);
            var requestOffset = Request.CreateRequest(builder, Command.Register, registerOffset.Value);
            builder.Finish(requestOffset.Value);
            SendFinishedMessage(builder);
        }

        private void SendFinishedMessage(FlatBufferBuilder finaliedBuilder)
        {
            var messageToSend = finaliedBuilder.DataBuffer.ToSizedArray();
            var messageSize = messageToSend.Length;
            sendMessageSize(messageSize);
            _stream.Write(messageToSend, 0, messageSize);
            _stream.Flush(); // Ensure data is sent immediately
        }

        public override String ToString()
        {
            return $"FbsClinet[{_host}:{_port} ({_priority})]";
        }

        protected override void SendRegistrationMessage()
        {
            var builder = new FlatBufferBuilder(64);
            var originOffset = builder.CreateString("HyperionScreenCap");
            var registerOffset = Register.CreateRegister(builder, originOffset, _priority);
            var requestOffset = Request.CreateRequest(builder, Command.Register, registerOffset.Value);
            builder.Finish(requestOffset.Value);
            SendFinishedMessage(builder);
        }

        public override void SendInitialFrame(int width, int height)
        {
            // Send a black frame to initialize the connection
            byte[] blackFrame = new byte[width * height * 3];
            SendImageDataMessage(blackFrame, width, height);
        }
    }
}
