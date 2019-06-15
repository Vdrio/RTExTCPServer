using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCPBindings
{
    public enum PacketType
    {
        ConnectionOk = 1,
        ClientSuccess = 2,
        StringMessage = 3,
        Range = 4,
        FragmentedRange = 5,
        SelectedRange = 6,

        //Border Color Changes
        TopBorderColor = 7,
        BottomBorderColor = 8,
        RightBorderColor = 9,
        LeftBorderColor = 10,
        InsideVerticalColor = 11,
        InsideHorizontalColor = 12,
        DiagonalUpColor = 13,
        DiagonalDownColor = 14,
        //Top Left Bottom Right
        BoxBorderColor = 15,
        //Top Left Bottom Right InsideHor InsideVert DiagonalUp DiagonalDown
        AllBorderColor = 16,
    }


}
