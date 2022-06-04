using EZSpreadsheet.StyleEnums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    class EZStyle
    {
        public EZFont Font { get; set; } = EZFont.Calibri;
        public EZFontColor FontColor { get; set; } = EZFontColor.Black;

        private uint _fontSize = 11;
        public uint FontSize
        {
            get => _fontSize;
            set
            {
                if ((value > 0) && (value <= 409))
                {
                    _fontSize = value;
                }
            }
        }

        private uint _numberFormat = 0;
        public uint NumberFormat
        {
            get => _numberFormat;
            set
            {
                if ((value >= 0) && (value <= 49))
                {
                    _numberFormat = value;
                }
            }
        }
    }
}
