using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;

namespace CloudyWing.Spreadsheet {

    [TypeConverter(typeof(VerticalAlignmentConverter))]
    public enum VerticalAlignment {
        Top = 0,
        Middle = 1,
        Bottom = 2
    }

    internal class VerticalAlignmentConverter : EnumConverter {

        private static Dictionary<string, VerticalAlignment> valueMap = new Dictionary<string, VerticalAlignment>(StringComparer.OrdinalIgnoreCase) {
            [nameof(VerticalAlignment.Top)] = VerticalAlignment.Top,
            [nameof(VerticalAlignment.Middle)] = VerticalAlignment.Middle,
            [nameof(VerticalAlignment.Bottom)] = VerticalAlignment.Bottom

        };

        public VerticalAlignmentConverter() : base(typeof(VerticalAlignment)) { }

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
                    return VerticalAlignment.Middle;
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
            if (destinationType == typeof(string) && valueMap.ContainsValue((VerticalAlignment)value)) {
                return value.ToString();
            }
            
            return base.ConvertTo(context, culture, value, destinationType);
        }
    }
}
