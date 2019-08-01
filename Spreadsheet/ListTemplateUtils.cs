using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace CloudyWing.Spreadsheet {
    public static class ListTemplateUtils {
        /// <summary>
        /// Config設定的格式、粗體、置中
        /// </summary>
        public static CellStyle HeaderStyle = CellStyle.CreateConfigStyle()
            .CloneAndSetHorizontalAlignment(HorizontalAlignment.Center)
            .CloneAndSetBorder(true)
            .CloneAndSetFont(HeaderFont);

        /// <summary>
        /// Config設定的格式、置左
        /// </summary>
        public static CellStyle TextStyle = CellStyle.CreateConfigStyle()
            .CloneAndSetBorder(true)
            .CloneAndSetHorizontalAlignment(HorizontalAlignment.Left);

        /// <summary>
        /// Config設定的格式、置右
        /// </summary>
        public static CellStyle NumberStyle = CellStyle.CreateConfigStyle()
            .CloneAndSetBorder(true)
            .CloneAndSetHorizontalAlignment(HorizontalAlignment.Right);

        /// <summary>
        /// Config設定的格式、置右
        /// </summary>
        public static CellStyle DateTimeStyle = CellStyle.CreateConfigStyle()
            .CloneAndSetBorder(true)
            .CloneAndSetHorizontalAlignment(HorizontalAlignment.Right);

        /// <summary>
        /// Config設定的格式、粗體
        /// </summary>
        public static CellFont HeaderFont = CellFont.CreateConfigFont()
            .CloneAndSetStyle(CellFont.CreateConfigFont().Style | FontStyles.IsBold);

        internal static IDictionary<string, object> ConvertToDictionary(object valueObj) {
            IDictionary<string, object> dic = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(valueObj);
            foreach (PropertyDescriptor prop in props) {
                object val = prop.GetValue(valueObj);
                dic.Add(prop.Name, val);
            }

            return dic;
        }
    }
}