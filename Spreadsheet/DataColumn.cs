using System;
using System.Collections.Generic;
using System.Drawing;

using static CloudyWing.Spreadsheet.ListTemplateUtils;

namespace CloudyWing.Spreadsheet {

    /// <summary>
    /// Excel匯出資料各標題欄位設定
    /// </summary>
    public class DataColumn<T> {

        private DataColumnCollection<T> childColumns;
        private Point point;

        /// <summary>
        /// 顯示文字
        /// </summary>
        public string HeaderText { get; set; }

        /// <summary>
        /// 對應DataSource的Property Name
        /// </summary>
        public string DataKey { get; set; }

        /// <summary>
        /// 標題的儲存格樣式
        /// </summary>
        public CellStyle HeaderStyle { get; set; }

        /// <summary>
        /// 資料儲存格樣式，若有設定ItemStyleFunctor，則無效果
        /// </summary>
        public CellStyle ItemStyle { get; set; }

        /// <summary>
        /// 資料儲存格樣式，優先權高於ItemStyle
        /// </summary>
        public Func<object, T, CellStyle> ItemStyleFunctor { get; set; }

        /// <summary>
        /// 修正顯示值的委派
        /// </summary>
        public Func<object, T, object> ContentRender { get; set; }

        /// <summary>
        /// 座標
        /// </summary>
        public Point Point {
            get {
                return point;
            }
            internal set {
                point = value;
                ChildColumns.ResetColumnsPoint(value + new Size(0, RowSpan));
            }
        }

        public int ColumnSpan => ChildColumns.Count == 0 ? 1 : ChildColumns.ColumnSpan;

        public int RowSpan => ParentColumns.RowSpan - ChildColumns.RowSpan;

        /// <summary>
        /// 自己和子層共幾層Column，用來計算RowSpan
        /// </summary>
        public int ColumnLayers => ChildColumns.Count == 0 ? 1 : ChildColumns.RowSpan + 1;

        /// <summary>
        /// 子標題欄位設定的集合
        /// </summary>
        public DataColumnCollection<T> ChildColumns {
            get {
                if (childColumns == null) {
                    childColumns = new DataColumnCollection<T>(this);
                }
                return childColumns;
            }
        }

        /// <summary>
        /// 自身標題欄位所屬的父集合
        /// </summary>
        public DataColumnCollection<T> ParentColumns { get; internal set; }

        internal virtual object GetContentValue(T valueObject) {
            IDictionary<string, object> data = ConvertToDictionary(valueObject);

            if (string.IsNullOrWhiteSpace(DataKey)) {
                return ContentRender == null ? "" : ContentRender(null, valueObject);
            }
            object value = data[DataKey];

            return ContentRender == null ? (value) : ContentRender(value, valueObject);
        }

        internal virtual CellStyle GetDataCellStyle(T valueObject) {
            IDictionary<string, object> data = ConvertToDictionary(valueObject);
            object value = string.IsNullOrWhiteSpace(DataKey) ? null : data[DataKey];

            return ItemStyleFunctor == null ? ItemStyle : ItemStyleFunctor(value, valueObject);
        }
    }

    public class DataColumn<TProperty, TObject> : DataColumn<TObject> {

        /// <summary>
        /// 資料儲存格樣式
        /// </summary>
        public new Func<TProperty, TObject, CellStyle> ItemStyleFunctor { get; set; }

        /// <summary>
        /// 修正顯示值的委派
        /// </summary>
        public new Func<TProperty, TObject, object> ContentRender { get; set; }

        internal override object GetContentValue(TObject valueObject) {
            IDictionary<string, object> data = ConvertToDictionary(valueObject);

            if (string.IsNullOrWhiteSpace(DataKey)) {
                return ContentRender == null ? "" : ContentRender(default(TProperty), valueObject);
            }
            object value = data[DataKey];

            if (ContentRender == null) {
                return value;
            }

            return ChangeValueTypeForFunc(ContentRender, value, valueObject);
        }

        internal override CellStyle GetDataCellStyle(TObject valueObject) {

            if (ItemStyleFunctor == null) {
                return ItemStyle;
            }

            if (string.IsNullOrWhiteSpace(DataKey)) {
                return ItemStyleFunctor(default(TProperty), valueObject);
            }

            IDictionary<string, object> data = ConvertToDictionary(valueObject);
            object value = data[DataKey];

            return ChangeValueTypeForFunc(ItemStyleFunctor, value, valueObject);
        }

        private TResult ChangeValueTypeForFunc<TResult>(Func<TProperty, TObject, TResult> func, object value, TObject valueObject) {

            if (value == null) {
                return func(default(TProperty), valueObject);
            }

            Type fromType = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();
            Type toType = Nullable.GetUnderlyingType(typeof(TProperty)) ?? typeof(TProperty);

            if (toType.IsPrimitive && typeof(IConvertible).IsAssignableFrom(fromType)) {
                return func((TProperty)Convert.ChangeType(value, toType), valueObject);
            }

            return func((TProperty)value, valueObject);
        }
    }
}