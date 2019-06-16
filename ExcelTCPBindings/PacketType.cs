using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTCPBindings
{
    public enum PacketType
    {
    //Systemic change enumerations
        //Changes.TCP Related (0s)
            ConnectionOk = 1,
            ClientSuccess = 2,
            StringMessage = 3,
            Range = 4,
            FragmentedRange = 5,
            SelectedRange = 6,

        //Changes.Cells (1000s) --------------o-----------------o-----------------o-----------------1000s-----------------o-----------------o-----------------o--------------
            //Changes.Cells.Border.Color *********************************************************************************************************************(1000-1049)
                //General Colors
                    DiagonalUpColor = 1000,                            //Target.Borders[XlBordersIndex.xlDiagonalUp].Color
                    DiagonalDownColor = 1001,                          //Target.Borders[XlBordersIndex.xlDiagonalDown].Color
                    BottomBorderColor = 1002,                          //Target.Borders[XlBordersIndex.xlEdgeBottom].Color
                    LeftBorderColor = 1003,                            //Target.Borders[XlBordersIndex.xlEdgeLeft].Color
                    RightBorderColor = 1004,                           //Target.Borders[XlBordersIndex.xlEdgeRight].Color
                    TopBorderColor = 1005,                             //Target.Borders[XlBordersIndex.xlEdgeTop].Color
                    InsideHorizontalColor = 1006,                      //Target.Borders[XlBordersIndex.xlInsideHorizontal].Color
                    InsideVerticalColor = 1007,                        //Target.Borders[XlBordersIndex.xlInsideVertical].Color
                    //Applies to Top Left Bottom Right
                    BoxBorderColor = 1008,                             //Target.Borders.Color   ??????                  
                    //Applies to Top Left Bottom Right InsideHor InsideVert DiagonalUp DiagonalDown
                    AllBorderColor = 1009,                             //Range.Borders.Color    ??????
                //Theme Colors
                    DiagonalUpThemeColor = 1020,                       //Target.Borders[XlBordersIndex.xlDiagonalUp].ThemeColor
                    DiagonalDownThemeColor = 1021,                     //Target.Borders[XlBordersIndex.xlDiagonalDown].ThemeColor
                    BottomBorderThemeColor = 1022,                     //Target.Borders[XlBordersIndex.xlEdgeBottom].ThemeColor
                    LeftBorderThemeColor = 1023,                       //Target.Borders[XlBordersIndex.xlEdgeLeft].ThemeColor
                    RightBorderThemeColor = 1024,                      //Target.Borders[XlBordersIndex.xlEdgeRight].ThemeColor
                    TopBorderThemeColor = 1025,                        //Target.Borders[XlBordersIndex.xlEdgeTop].ThemeColor
                    InsideHorizontalThemeColor = 1026,                 //Target.Borders[XlBordersIndex.xlInsideHorizontal].ThemeColor
                    InsideVerticalThemeColor = 1027,                   //Target.Borders[XlBordersIndex.xlInsideVertical].ThemeColor
                    //Applies to Top Left Bottom Right
                    BoxBorderThemeColor = 1028,                        //Target.Borders.ThemeColor     ??????                
                    //Applies to Top Left Bottom Right InsideHor InsideVert DiagonalUp DiagonalDown
                    AllBorderThemeColor = 1029,                        //Range.Borders.ThemeColor    ??????
                //Tint and Shade
                    DiagonalUpTintShade = 1040,                        //Target.Borders[XlBordersIndex.xlDiagonalUp].TintAndShade
                    DiagonalDownTintShade = 1041,                      //Target.Borders[XlBordersIndex.xlDiagonalDown].TintAndShade
                    BottomBorderTintShade = 1042,                      //Target.Borders[XlBordersIndex.xlEdgeBottom].TintAndShade
                    LeftBorderTintShade = 1043,                        //Target.Borders[XlBordersIndex.xlEdgeLeft].TintAndShade
                    RightBorderTintShade = 1044,                       //Target.Borders[XlBordersIndex.xlEdgeRight].TintAndShade
                    TopBorderTintShade = 1045,                         //Target.Borders[XlBordersIndex.xlEdgeTop].TintAndShade
                    InsideHorizontalTintShade = 1046,                  //Target.Borders[XlBordersIndex.xlInsideHorizontal].TintAndShade
                    InsideVerticalTintShade = 1047,                    //Target.Borders[XlBordersIndex.xlInsideVertical].TintAndShade
                    //Applies to Top Left Bottom Right
                    BoxBorderTintShade = 1048,                         //Target.Borders.TintAndShade     ??????                
                    //Applies to Top Left Bottom Right InsideHor InsideVert DiagonalUp DiagonalDown
                    AllBorderTintShade = 1049,                         //Range.Borders.TintAndShade      ??????

            //Changes.Cells.Border.Weight ********************************************************************************************************************(1070-1099)
                DiagonalUpWeight = 1070,                               //Target.Borders[XlBordersIndex.xlDiagonalUp].Weight
                DiagonalDownWeight = 1071,                             //Target.Borders[XlBordersIndex.xlDiagonalDown].Weight
                BottomBorderWeight = 1072,                             //Target.Borders[XlBordersIndex.xlEdgeBottom].Weight
                LeftBorderWeight = 1073,                               //Target.Borders[XlBordersIndex.xlEdgeLeft].Weight
                RightBorderWeight = 1074,                              //Target.Borders[XlBordersIndex.xlEdgeRight].Weight
                TopBorderWeight = 1075,                                //Target.Borders[XlBordersIndex.xlEdgeTop].Weight
                InsideHorizontalWeight = 1076,                         //Target.Borders[XlBordersIndex.xlInsideHorizontal].Weight
                InsideVerticalWeight = 1057,                           //Target.Borders[XlBordersIndex.xlInsideVertical].Weight
                //Applies to Top Left Bottom Right
                BoxBorderWeight = 1078,                                //Target.Borders.Weight     ??????                
                //Applies to Top Left Bottom Right InsideHor InsideVert DiagonalUp DiagonalDown
                AllBorderWeight = 1079,                                //Range.Borders.Weight      ??????

            //Changes.Cells.Border.LineStyle *****************************************************************************************************************(1100-1199)
                DiagonalUpStyle = 1100,                                //Target.Borders[XlBordersIndex.xlDiagonalUp].LineStyle
                DiagonalDownLineStyle = 1101,                          //Target.Borders[XlBordersIndex.xlDiagonalDown].LineStyle
                BottomBorderLineStyle = 1102,                          //Target.Borders[XlBordersIndex.xlEdgeBottom].LineStyle
                LeftBorderLineStyle = 1103,                            //Target.Borders[XlBordersIndex.xlEdgeLeft].LineStyle
                RightBorderLineStyle = 1104,                           //Target.Borders[XlBordersIndex.xlEdgeRight].LineStyle
                TopBorderLineStyle = 1105,                             //Target.Borders[XlBordersIndex.xlEdgeTop].LineStyle
                InsideHorizontalLineStyle = 1106,                      //Target.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle
                InsideVerticalLineStyle = 1107,                        //Target.Borders[XlBordersIndex.xlInsideVertical].LineStyle
                //Applies to Top Left Bottom Right
                BoxBorderLineStyle = 1108,                             //Target.Borders.LineStyle     ??????                
                //Applies to Top Left Bottom Right InsideHor InsideVert DiagonalUp DiagonalDown
                AllBorderLineStyle = 1109,                             //Range.Borders.LineStyle      ??????

        //range.interior

        //Changes.Cells.Number (i.e. general, number, currency as seen in format cells) **********************************************************************(1200-1299)

        //Changes.Cells.Allignment (i.e. Text allignment as seen in format cells) ****************************************************************************(1300-1399)

        //Changes.Cells.Font (i.e. Text Font Type as seen in format cells) ***********************************************************************************(1400-1499)

        //Changes.Cells.Fill (i.e. Text Fill as seen in format cells) ****************************************************************************************(1500-1599)

        //Changes.Cells.Protection (i.e. Cell Protection as seen in format cells) ****************************************************************************(1600-1699)

        //Changes.Cells.ConditionalFormating *****************************************************************************************************************(1600-1699)

        //Changes.Insert (2000s) -------------o-----------------o-----------------o-----------------2000s-----------------o-----------------o-----------------o--------------
        //Changes.PageLayout (3000s) ---------o-----------------o-----------------o-----------------3000s-----------------o-----------------o-----------------o--------------
        //Changes.Tables (4000s) -------------o-----------------o-----------------o-----------------4000s-----------------o-----------------o-----------------o--------------
        //Changes.Charts (5000s) -------------o-----------------o-----------------o-----------------5000s-----------------o-----------------o-----------------o--------------
        //Changes.Sheets (6000s) -------------o-----------------o-----------------o-----------------6000s-----------------o-----------------o-----------------o--------------
        //Changes.Review (7000s) -------------o-----------------o-----------------o-----------------7000s-----------------o-----------------o-----------------o--------------
        //Changes.Review.Comments ****************************************************************************************************************************(7000-7099)

        //Changes.View (8000s) ---------------o-----------------o-----------------o-----------------8000s-----------------o-----------------o-----------------o--------------
        
                //For freezing panes
        //https://stackoverflow.com/questions/5060488/how-to-freeze-top-row-and-apply-filter-in-excel-automation-with-c-sharp
    }


}
