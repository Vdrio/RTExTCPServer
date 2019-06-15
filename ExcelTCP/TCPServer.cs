using ExcelTCPBindings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCP
{
    public class TCPServer
    {
        private static Socket ServerSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        private static byte[] Buffer = new byte[1024];
        public static Client[] Clients = new Client[Constants.MAX_CONNECTIONS];

        public static void SetupServer()
        {
            for (int i = 0; i < Constants.MAX_CONNECTIONS; i++)
            {
                Clients[i] = new Client();
            }
            ServerSocket.Bind(new IPEndPoint(IPAddress.Any, 8081));
            ServerSocket.Listen(10);
            ServerSocket.BeginAccept(new AsyncCallback(AcceptCallback), null);
            NetworkDataHandler.InitializeNetworkPackages();
            NetworkDataHandler.SelectionReceived += NetworkDataHandler_SelectionReceived;
        }

        private static void NetworkDataHandler_SelectionReceived(object sender, EventArgs e)
        {
            Tuple<int, byte[]> info = (Tuple<int, byte[]>)sender;
            for (int i = 0; i < Clients.Count(); i++)
            {
                if (i == info.Item1)
                {
                    continue;
                }

                SendDataTo(i, info.Item2);
            }
        }

        private static void AcceptCallback(IAsyncResult ar)
        {
            Socket socket = ServerSocket.EndAccept(ar);
            ServerSocket.BeginAccept(new AsyncCallback(AcceptCallback), null);

            for (int i = 0; i < Constants.MAX_CONNECTIONS; i++)
            {
                if (Clients[i].socket == null)
                {
                    Clients[i].socket = socket;
                    Clients[i].index = i;
                    Clients[i].ip = socket.RemoteEndPoint.ToString();
                    Clients[i].StartClient();
                    Console.WriteLine(string.Format("Connection from '{0}' received.", Clients[i].ip));
                    SendConnectionOK(i);
                    return;
                }
            }
        }

        public static void SendDataTo(int index, byte[] data)
        {
            byte[] sizeInfo = new byte[4];
            sizeInfo[0] = (byte)data.Length;
            sizeInfo[1] = (byte)(data.Length >> 8);
            sizeInfo[2] = (byte)(data.Length >> 16);
            sizeInfo[3] = (byte)(data.Length >> 24);

            Clients[index].socket.Send(sizeInfo);
            Clients[index].socket.Send(data);
        }

        public static void SendConnectionOK(int index)
        {
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteInteger(1);
            buffer.WriteString("Welcome to R-Tex");
            SendDataTo(index, buffer.ToArray());
            buffer.Dispose();
        }

    }

    public class Client
    {
        public int index;
        public string ip;
        public Socket socket;
        public bool closing = false;
        private byte[] Buffer = new byte[1024];

        public void StartClient()
        {
            socket.BeginReceive(Buffer, 0, Buffer.Length, SocketFlags.None, new AsyncCallback(ReceiveCallback), socket);
        }

        private void ReceiveCallback(IAsyncResult ar)
        {
            Socket socket = (Socket)ar.AsyncState;

            try
            {
                int received = socket.EndReceive(ar);
                if (received <= 0)
                {
                    CloseClient(index);
                }
                else
                {
                    byte[] dataBuffer = new byte[received];
                    Array.Copy(Buffer, dataBuffer, received);
                    NetworkDataHandler.HandleNetworkInformation(index, dataBuffer);
                    socket.BeginReceive(Buffer, 0, Buffer.Length, SocketFlags.None, new AsyncCallback(ReceiveCallback), socket);
                }
            }
            catch (Exception ex)
            {
                CloseClient(index);
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        private void CloseClient(int ind)
        {
            closing = true;
            Console.WriteLine(string.Format("Connection from {0} has been terminated.", ip));
            socket.Close();
            socket = null;
        }
    }
}
