using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCPBindings
{
    [Serializable]
    public class PacketBuffer : IDisposable
    {
        List<byte> bufferList;
        byte[] readBuffer;
        int readPosition;
        bool buffUpdate = false;

        public PacketBuffer()
        {
            bufferList = new List<byte>();
            readPosition = 0;
        }

        public int GetReadPosition()
        {
            return readPosition;
        }

        public byte[] ToArray()
        {
            return bufferList.ToArray();
        }

        public int Count()
        {
            return bufferList.Count;
        }

        public int Length()
        {
            return Count() - readPosition;
        }

        public void Clear()
        {
            bufferList.Clear();
            readPosition = 0;

        }

        //Write Data
        public void WriteBytes(byte[] input)
        {
            bufferList.AddRange(input);
            buffUpdate = true;
        }

        public void WriteByte(byte input)
        {
            bufferList.Add(input);
            buffUpdate = true;
        }

        public void WriteInteger(int input)
        {
            bufferList.AddRange(BitConverter.GetBytes(input));
            buffUpdate = true;
        }
        public void WriteFloat(float input)
        {
            bufferList.AddRange(BitConverter.GetBytes(input));
            buffUpdate = true;
        }
        public void WriteString(string input)
        {
            bufferList.AddRange(BitConverter.GetBytes(input.Length));
            bufferList.AddRange(Encoding.ASCII.GetBytes(input));
            buffUpdate = true;
        }

        public void WriteSelectedRange(RangePacket input)
        {
            try
            {
                Debug.WriteLine("Writing selected range:");
                BinaryFormatter bf = new BinaryFormatter();
                byte[] data;
                using (var ms = new MemoryStream())
                {
                    bf.Serialize(ms, input);
                    data = new byte[ms.Length];
                    data = ms.ToArray();
                }
                Debug.WriteLine(data.Length);
                bufferList.AddRange(BitConverter.GetBytes(data.Length));
                bufferList.AddRange(data);
                buffUpdate = true;
            }
            catch(Exception ex) { Debug.WriteLine("Writing selected range: "  + ex); }
        }

        //Read Data
        public int ReadInteger(bool peek = true)
        {
            if (bufferList.Count > readPosition)
            {
                if (buffUpdate)
                {
                    readBuffer = bufferList.ToArray();
                    buffUpdate = false;
                }

                int value = BitConverter.ToInt32(readBuffer, readPosition);
                if (peek && bufferList.Count > readPosition)
                {
                    readPosition += 4;
                }
                return value;
            }
            else
            {
                Console.WriteLine("Buffer list past limit");
                return -1;
            }
        }
        public float ReadFloat(bool peek = true)
        {
            if (bufferList.Count > readPosition)
            {
                if (buffUpdate)
                {
                    readBuffer = bufferList.ToArray();
                    buffUpdate = false;
                }

                float value = BitConverter.ToSingle(readBuffer, readPosition);
                if (peek && bufferList.Count > readPosition)
                {
                    readPosition += 4;
                }
                return value;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Buffer list past limit");
                return -1;
            }
        }
        public byte ReadByte(bool peek = true)
        {
            if (bufferList.Count > readPosition)
            {
                if (buffUpdate)
                {
                    readBuffer = bufferList.ToArray();
                    buffUpdate = false;
                }

                byte value = readBuffer[readPosition];
                if (peek && bufferList.Count > readPosition)
                {
                    readPosition += 1;
                }
                return value;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Buffer list past limit");
                return new byte();
            }
        }
        public byte[] ReadBytes(int length, bool peek = true)
        {
            if (buffUpdate)
            {
                readBuffer = bufferList.ToArray();
                buffUpdate = false;
            }

            byte[] value = bufferList.GetRange(readPosition, length).ToArray();
            if (peek && bufferList.Count > readPosition)
            {
                readPosition += length;
            }
            return value;

        }
        public string ReadString(bool peek = true)
        {
            int length = ReadInteger();
            if (buffUpdate)
            {
                readBuffer = bufferList.ToArray();
                buffUpdate = false;
            }

            string value = Encoding.ASCII.GetString(readBuffer, readPosition, length);
            if (peek && bufferList.Count > readPosition)
            {
                readPosition += length;
            }
            return value;

        }

        public RangePacket ReadRange(bool peek = true)
        {
            int length = ReadInteger();
            if (buffUpdate)
            {
                readBuffer = bufferList.ToArray();
                buffUpdate = false;
            }

            BinaryFormatter b = new BinaryFormatter();
            byte[] rangeBytes = new byte[length];
            Array.Copy(readBuffer, readPosition, rangeBytes, 0, length);
            Debug.WriteLine("copied array to read range bytes");
            RangePacket val = (RangePacket)b.Deserialize(new MemoryStream(rangeBytes));

            if (peek && bufferList.Count > readPosition)
            {
                readPosition += length;
            }
            return val;

        }

        public RangePacket ReadSelectedRange(bool peek = true)
        {
            int length = ReadInteger();
            if (buffUpdate)
            {
                readBuffer = bufferList.ToArray();
                buffUpdate = false;
            }

            BinaryFormatter b = new BinaryFormatter();
            byte[] rangeBytes = new byte[length];
            Array.Copy(readBuffer, readPosition, rangeBytes, 0, length);

            RangePacket val = (RangePacket)b.Deserialize(new MemoryStream(rangeBytes));

            if (peek && bufferList.Count > readPosition)
            {
                readPosition += length;
            }
            return val;
        }

        //IDisposable
        private bool disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    bufferList.Clear();
                }
                readPosition = 0;
            }
            disposed = true;

        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

    }
}
