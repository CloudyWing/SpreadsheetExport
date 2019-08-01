using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;

namespace CloudyWing.Spreadsheet {
    [TypeConverter(typeof(HorizontalAlignmentConverter))]
    public enum HorizontalAlignment {
        None = 0,
        Left = 1,
        Center = 2,
        Right = 3,
        Justify = 4
    }

    internal class HorizontalAlignmentConverter : EnumConverter {
        private static readonly Dictionary<string, HorizontalAlignment> valueMap = new Dictionary<string, HorizontalAlignment>(StringComparer.OrdinalIgnoreCase) {
            [nameof(HorizontalAlignment.None)] = HorizontalAlignment.None,
            [nameof(HorizontalAlignment.Center)] = HorizontalAlignment.Center,
            [nameof(HorizontalAlignment.Right)] = HorizontalAlignment.Right,
            [nameof(HorizontalAlignment.Justify)] = HorizontalAlignment.Justify

        };

        public HorizontalAlignmentConverter() : base(typeof(HorizontalAlignment)) { }

        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType) {
            return sourceType == typeof(string) ? true : base.CanConvertFrom(context, sourceType);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value) {
            if (value == null) {
                return null;
            }

            if (value is string) {
                string textValue = ((string)value).Trim();
                if (textValue.Length == 0) {
                    return HorizontalAlignment.None;
                }

                if (valueMap.ContainsKey(textValue)) {
                    return valueMap[textValue];
                }
            }
            return base.ConvertFrom(context, culture, value);
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type sourceType) {
            if (sourceType == typeof(string)) {
                return true;
            }

            return base.CanConvertTo(context, sourceType);
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType) {
            if (destinationType == typeof(string) && valueMap.ContainsValue((HorizontalAlignment)value)) {
                return value.ToString();
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }
    }
}