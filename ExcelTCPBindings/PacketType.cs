using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCPBindings
{
    public enum PacketType
    {
        ConnectionOk = 1, ClientSuccess = 5, StringMessage = 2, Range = 3, FragmentedRange = 4, SelectedRange = 6
    }


}
