using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCPBindings
{
    [Serializable]
    public class RangePacket:ISerializable
    {
        public string RangeInfo { get; set; }
        public ExcelUser User { get; set; }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info?.AddValue("RangeInfo", RangeInfo);
            info?.AddValue("User", User);
        }

        public RangePacket()
        {

        }

        protected RangePacket(SerializationInfo info, StreamingContext context)
        {
            RangeInfo = (string)info.GetValue("RangeInfo", typeof(string));
            User = (ExcelUser)info.GetValue("User", typeof(ExcelUser));
        }
    }
}
