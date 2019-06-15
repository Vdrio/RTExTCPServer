using ExcelTCPBindings;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCPBindings
{
    public class NetworkDataHandler
    {
        private delegate void PacketData(int index, byte[] data);
        private static Dictionary<int, PacketData> Packets;

        public static event EventHandler SelectionReceived;
        public static event EventHandler ThankYouServer;
        public static event EventHandler SendSelectionUpdateToClients;

        static bool Server;

        public static void InitializeNetworkPackages(bool server = false)
        {
            Console.WriteLine("Initializing network packages...");
            Server = server;
            Packets = new Dictionary<int, PacketData>
            {
                { (int)PacketType.ConnectionOk, HandleConnectionOK }, { (int)PacketType.StringMessage, HandleStringMessage}
                ,{(int)PacketType.Range, HandleRangeUpdate}, {(int)PacketType.ClientSuccess, HandleThankYou }
                , { (int)PacketType.SelectedRange, HandleSelectionUpdate} 
            };
        }

        public static void HandleNetworkInformation(int index, byte[] data)
        {
            int packetNum;
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteBytes(data);
            packetNum = buffer.ReadInteger();
            Debug.WriteLine("Handling type: " + packetNum);
            buffer.Dispose();
            if (Packets.TryGetValue(packetNum, out PacketData Packet))
            {
                Packet.Invoke(index, data);
            }
        }

        private static void HandleThankYou(int index, byte[] data)
        {
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteBytes(data);
            buffer.ReadInteger();
            string msg = buffer.ReadString();
            buffer.Dispose();
            Console.WriteLine(msg);
            System.Diagnostics.Debug.WriteLine(string.Format("From {0}: {1}", index, msg));
        }

        private static void HandleConnectionOK(int index, byte[] data)
        {
            Debug.WriteLine("Handling Connection Ok");
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteBytes(data);
            buffer.ReadInteger();
            string msg = buffer.ReadString();
            buffer.Dispose();
            if (msg == null)
            {
                return;
            }
            ThankYouServer?.Invoke(msg, EventArgs.Empty);
            Debug.WriteLine(msg);
        }

        private static void HandleStringMessage(int index, byte[] data)
        {
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteBytes(data);
            buffer.ReadInteger();
            string msg = buffer.ReadString();
            buffer.Dispose();
            Console.WriteLine(msg);
        }

        private static void HandleSelectionUpdate(int index, byte[] data)
        {
            Console.WriteLine("Received selection update");
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteBytes(data);
            buffer.ReadInteger();
            RangePacket r = buffer.ReadSelectedRange();
            if (Server)
            {
                SelectionReceived?.Invoke(new Tuple<int, byte[]>(index, data), EventArgs.Empty);
            }
            else
            {
                SelectionReceived?.Invoke(r, EventArgs.Empty);
            }
            buffer.Dispose();
            Console.WriteLine(r.User.ToString() + "has updated their selection to: " + r.RangeInfo);
        }

        private static void HandleRangeUpdate(int index, byte[] data)
        {
            PacketBuffer buffer = new PacketBuffer();
            buffer.WriteBytes(data);
            buffer.ReadInteger();
            RangePacket r = buffer.ReadRange();
            buffer.Dispose();
            Console.WriteLine(r);
        }
    }
}
