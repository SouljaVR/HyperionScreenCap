﻿using System;
using Google.ProtocolBuffers;
using proto;
using static Humanizer.In;

namespace HyperionScreenCap.Networking
{
    class ProtoClient : HyperionClient
    {
        public ProtoClient(string host, int port, int priority, int messageDuration): base(host, port, priority, messageDuration)
        {
    
        }

        protected override void SendImageDataMessage(byte[] pixeldata, int width, int height)
        {
            var imageRequest = ImageRequest.CreateBuilder()
                .SetImagedata(ByteString.CopyFrom(pixeldata))
                .SetImageheight(height)
                .SetImagewidth(width)
                .SetPriority(_priority)
                .SetDuration(_messageDuration)
                .Build();

            var request = HyperionRequest.CreateBuilder()
                .SetCommand(HyperionRequest.Types.Command.IMAGE)
                .SetExtension(ImageRequest.ImageRequest_, imageRequest)
                .Build();

            SendRequest(request);
        }

        protected override void SendClearPriorityMessage()
        {
            var clearRequest = ClearRequest.CreateBuilder()
                .SetPriority(_priority)
                .Build();

            var request = HyperionRequest.CreateBuilder()
                .SetCommand(HyperionRequest.Types.Command.CLEAR)
                .SetExtension(ClearRequest.ClearRequest_, clearRequest)
                .Build();

            SendRequest(request);
        }

        private void SendRequest(IMessageLite request)
        {
            var size = request.SerializedSize;
            var header = new byte[4];
            header[0] = (byte)((size >> 24) & 0xFF);
            header[1] = (byte)((size >> 16) & 0xFF);
            header[2] = (byte)((size >> 8) & 0xFF);
            header[3] = (byte)(size & 0xFF);

            var headerSize = header.Length;
            _stream.Write(header, 0, headerSize);
            request.WriteTo(_stream);
            _stream.Flush();

            // Enable reply message if needed (debugging only).
            //var reply = ReceiveReply();
            //Console.WriteLine($@"Reply: {reply.ToString()}");
        }

        private HyperionReply ReceiveReply()
        {
            var header = new byte[4];
            _stream.Read(header, 0, 4);
            var size = (header[0] << 24) | (header[1] << 16) | (header[2] << 8) | (header[3]);
            var data = new byte[size];
            _stream.Read(data, 0, size);
            var reply = HyperionReply.ParseFrom(data);

            return reply;
        }

        public override String ToString()
        {
            return $"ProtoClient[{_host}:{_port} ({_priority})]";
        }

        protected override void SendRegistrationMessage()
        {
            // Assuming there's no specific RegisterRequest, we'll use the general HyperionRequest
            var request = proto.HyperionRequest.CreateBuilder()
                .SetCommand(proto.HyperionRequest.Types.Command.COLOR) // Use an existing command
                .Build();

            SendRequest(request);
        }

        public override void SendInitialFrame(int width, int height)
        {
            // Send a black frame to initialize the connection
            byte[] blackFrame = new byte[width * height * 3];
            SendImageDataMessage(blackFrame, width, height);
        }
    }
}
