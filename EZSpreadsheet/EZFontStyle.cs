using EZSpreadsheet.StyleEnums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    public class EZFontStyle: IEquatable<EZFontStyle>
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

        public bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        public bool IsUnderlined { get; set; } = false;

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

        public bool Equals(EZFontStyle? other)
        {
            if (other == null)
                return false;

            if (Font != other.Font)
                return false;

            if (FontColor != other.FontColor)
                return false;

            if (FontSize != other.FontSize)
                return false;

            if (FontColor != other.FontColor)
                return false;

            if (IsBold != other.IsBold || IsItalic != other.IsItalic || IsUnderlined != other.IsUnderlined)
                return false;

            return true;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as EZFontStyle);
        }
    }
}
