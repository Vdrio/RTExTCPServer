using ExcelTCPBindings;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCP
{
    public class TCPClient
    {

        public delegate void StringMessageReceived(string data);
        public delegate void ActionStatusUpdate(string data);
        public static StringMessageReceived OnStringMessageReceived;
        private static Socket ClientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        private byte[] AsyncBuffer = new byte[1024];


        public static void ConnectToServer()
        {
            Console.WriteLine("Connecting to server...");
            NetworkDataHandler.InitializeNetworkPackages();
            ClientSocket.BeginConnect("192.168.1.101", 8081, new AsyncCallback(ConnectCallback), ClientSocket);
        }

        public static void ConnectToServer(string ip, int port)
        {
            Console.WriteLine("Connecting to server...");
            ClientSocket.BeginConnect(ip, port, new AsyncCallback(ConnectCallback), ClientSocket);
        }

        private static void ConnectCallback(IAsyncResult ar)
        {
            try
            {
                ClientSocket.EndConnect(ar);
                while (true)
                {
                    OnReceive();
                }
            }
            catch { }
        }

        public static void EndConnection()
        {
            ClientSocket.Close();
            ClientSocket.Dispose();
        }

        private static void OnReceive()
        {
            byte[] sizeInfo = new byte[4];
            byte[] receivedBuffer = new byte[1024];
            int totalRead = 0, currentRead = 0;

            try
            {
                currentRead = totalRead = ClientSocket.Receive(sizeInfo);
                if (totalRead <= 0)
                {
                    Console.WriteLine("Not connected to server.");
                }
                else
                {
                    Console.WriteLine("Reading data...");
                    while (totalRead < sizeInfo.Length && currentRead > 0)
                    {
                        currentRead = ClientSocket.Receive(sizeInfo, totalRead, sizeInfo.Length - totalRead, SocketFlags.None);
                        totalRead += currentRead;
                    }

                    int messageSize = 0;
                    messageSize |= sizeInfo[0];
                    messageSize |= (sizeInfo[1] << 8);
                    messageSize |= (sizeInfo[2] << 16);
                    messageSize |= (sizeInfo[3] << 24);

                    byte[] data = new byte[messageSize];
                    totalRead = 0;
                    currentRead = totalRead = ClientSocket.Receive(data, totalRead, data.Length - totalRead, SocketFlags.None);
                    while (totalRead < messageSize && currentRead > 0)
                    {
                        currentRead = ClientSocket.Receive(data, totalRead, data.Length - totalRead, SocketFlags.None);
                        totalRead += currentRead;
                    }

                    Debug.WriteLine("Received message, handling");
                    NetworkDataHandler.HandleNetworkInformation(messageSize, data);
                    

                }
            }
            catch
            {
                Console.WriteLine("Not connected to server.");
            }
        }

        public static void SendData(byte[] data)
        {
            Debug.WriteLine("Sending data");
            ClientSocket.Send(data);
            Debug.WriteLine("Sent data");
        }

        public static void ThankYouServer()
        {
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteInteger((int)PacketType.ClientSuccess);
            buffer.WriteString("Thank you for letting me connect");
            SendData(buffer.ToArray());
            buffer.Dispose();
        }

        public static void SendSelectionUpdate(RangePacket packet)
        {
           // ThankYouServer();
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteInteger((int)PacketType.SelectedRange);
            buffer.WriteSelectedRange(packet);
            Debug.WriteLine("Buffer created: " + buffer.Length());
            SendData(buffer.ToArray());
            buffer.Dispose();
        }
    }
}
