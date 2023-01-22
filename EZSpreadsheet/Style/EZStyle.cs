using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet.Style
{
    public class EZStyle : IEquatable<EZStyle>
    {
        public EZFont Font { get; set; } = EZFont.Calibri;
        public EZColor FontColor { get; set; } = EZColor.Black;

        private uint _fontSize = 11;
        public uint FontSize
        {
            get => _fontSize;
            set
            {
                if (value > 0 && value <= 409)
                {
                    _fontSize = value;
                }
            }
        }

        public bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        public bool IsUnderlined { get; set; } = false;

        public EZColor BorderColor { get; set; } = EZColor.Black;

        public EZBorder BorderType { get; set; } = EZBorder.None;

        public EZColor FillColor { get; set; } = EZColor.White;

        private uint _numberFormatId = 0;
        public uint NumberFormatId
        {
            get => _numberFormatId;
            set
            {
                if (value >= 0 && value <= 49)
                {
                    _numberFormatId = value;
                }
            }
        }

        internal uint FontId { get; set; }
        internal uint BorderId { get; set; }
        internal uint FillId { get; set; }

        internal bool FontEquals(EZStyle? other)
        {
            if (other == null)
                return false;

            if (Font != other.Font || FontColor != other.FontColor || FontSize != other.FontSize)
                return false;

            if (IsBold != other.IsBold || IsItalic != other.IsItalic || IsUnderlined != other.IsUnderlined)
                return false;

            return true;
        }

        internal bool BorderEquals(EZStyle? other)
        {
            if (other == null)
                return false;

            if (BorderColor != other.BorderColor || BorderType != other.BorderType)
                return false;

            return true;
        }

        internal bool FillEquals(EZStyle? other)
        {
            if (other == null)
                return false;

            if (FillColor != other.FillColor)
                return false;

            return true;
        }

        public bool Equals(EZStyle? other)
        {
            if (other == null)
                return false;

            if (Font != other.Font || FontColor != other.FontColor || FontSize != other.FontSize)
                return false;

            if (IsBold != other.IsBold || IsItalic != other.IsItalic || IsUnderlined != other.IsUnderlined)
                return false;

            if (BorderColor != other.BorderColor || BorderType != other.BorderType)
                return false;

            if (FillColor != other.FillColor)
                return false;

            if (NumberFormatId != other.NumberFormatId)
                return false;

            return true;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as EZStyle);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FontId, BorderId, FillId, NumberFormatId);
        }
    }
}
