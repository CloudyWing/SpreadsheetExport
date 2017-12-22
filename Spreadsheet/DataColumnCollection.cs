using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using static CloudyWing.Spreadsheet.ListTemplateUtils;

namespace CloudyWing.Spreadsheet {

    public class DataColumnCollection<T> : Collection<DataColumn<T>> {

        private DataColumn<T> parentItem;

        internal DataColumnCollection(DataColumn<T> cell) {
            parentItem = cell;
        }

        /// <summary>
        /// 行跨度
        /// </summary>
        public int ColumnSpan => Count == 0 ? 0 : this.Sum(x => x.ColumnSpan);

        /// <summary>
        /// 列跨度
        /// </summary>
        public int RowSpan => Count == 0 ? 0 : this.Max(x => x.ColumnLayers);

        /// <summary>
        /// 根資料列的集合
        /// </summary>
        public DataColumnCollection<T> RootColumns {
            get {
                DataColumnCollection<T> items = this;
                while (items.parentItem != null && items.parentItem.ParentColumns != null) {
                    items = items.parentItem.ParentColumns;
                }
                return items;
            }
        }

        /// <summary>
        /// 從根節點往下重設座標
        /// </summary>
        internal void ResetRootPoint() {
            RootColumns.ResetColumnsPoint(Point.Empty);
        }

        /// <summary>
        /// 重設底下所有DataColumn的座標
        /// </summary>
        /// <param name="point"></param>
        internal void ResetColumnsPoint(Point point) {
            Size offset = new Size();

            foreach (DataColumn<T> item in this) {
                item.Point = point + offset;
                offset.Width = offset.Width + item.ColumnSpan;
            }
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void Add(
            string headerText, string dataKey = "",
            Func<object, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            Add(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void Add(
            string headerText, Expression<Func<T, object>> expression,
            Func<object, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            Add(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        private string GetDataKeyByExpression(Expression<Func<T, object>> expression) {
            if (expression == null) {
                throw new ArgumentNullException(nameof(expression));
            }

            // 如果是Value Type 的話Body會是UnaryExpression
            // Reference Type才會是直接取得到MemberExpression
            UnaryExpression unary = expression.Body as UnaryExpression;
            MemberExpression member = expression.Body as MemberExpression ??
                (unary != null ? unary.Operand as MemberExpression : null);

            if (member == null) {
                throw new ArgumentException($"Expression格式錯誤。", nameof(expression));
            }
            return member.Member.Name;
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleFunc">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void Add(
            string headerText, string dataKey,
            Func<object, T, object> render,
            CellStyle? headerStyle, Func<object, T, CellStyle> dataStyleFunc
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            Add(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleFunc
            });
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleFunc">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void Add(
            string headerText, Expression<Func<T, object>> expression,
            Func<object, T, object> render,
            CellStyle? headerStyle, Func<object, T, CellStyle> dataStyleFunc
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            Add(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleFunc
            });
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void Add<TProperty>(
            string headerText, string dataKey = "",
            Func<TProperty, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            Add(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void Add<TProperty>(
            string headerText, Expression<Func<T, object>> expression,
            Func<TProperty, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            Add(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleExpression">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void Add<TProperty>(
            string headerText, string dataKey,
            Func<TProperty, T, object> render,
            CellStyle? headerStyle, Func<TProperty, T, CellStyle> dataStyleExpression
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            Add(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleExpression
            });
        }

        /// <summary>
        /// 增加一筆DataColumn
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleExpression">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void Add<TProperty>(
            string headerText, Expression<Func<T, object>> expression,
            Func<TProperty, T, object> render,
            CellStyle? headerStyle, Func<TProperty, T, CellStyle> dataStyleExpression
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            Add(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleExpression
            });
        }

        /// <exception cref="NullReferenceException">尚未建立任何DataColumn<T>。</exception>
        public void AddChildToLast(DataColumn<T> childHeader) {
            DataColumn<T> header = this.LastOrDefault();
            if (header == null) {
                throw new NullReferenceException($"尚未建立任何{nameof(DataColumn<T>)}。");
            }

            header.ChildColumns.Add(childHeader);
        }

        /// <summary>
        /// 增加一筆子DataHeader至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void AddChildToLast(
            string headerText, string dataKey = "",
            Func<object, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            AddChildToLast(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        /// <summary>
        /// 增加一筆子DataHeader至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void AddChildToLast(
            string headerText, Expression<Func<T, object>> expression,
            Func<object, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            AddChildToLast(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        /// <summary>
        /// 增加一筆子DataColumn至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleExpression">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void AddChildToLast(
            string headerText, string dataKey,
            Func<object, T, object> render,
            CellStyle? headerStyle, Func<object, T, CellStyle> dataStyleExpression
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            AddChildToLast(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleExpression
            });
        }

        /// <summary>
        /// 增加一筆子DataColumn至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleExpression">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void AddChildToLast(
            string headerText, Expression<Func<T, object>> expression,
            Func<object, T, object> render,
            CellStyle? headerStyle, Func<object, T, CellStyle> dataStyleExpression
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            AddChildToLast(new DataColumn<T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleExpression
            });
        }

        /// <summary>
        /// 增加一筆子DataHeader至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void AddChildToLast<TProperty>(
            string headerText, string dataKey = "",
            Func<TProperty, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            AddChildToLast(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        /// <summary>
        /// 增加一筆子DataHeader至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="itemStyle">資料的儲存格式，預設ListTemplateUtils.TextStyle</param>
        public void AddChildToLast<TProperty>(
            string headerText, Expression<Func<T, object>> expression,
            Func<TProperty, T, object> render = null,
            CellStyle? headerStyle = null, CellStyle? itemStyle = null
        ) {
            headerStyle = headerStyle ?? HeaderStyle;
            itemStyle = itemStyle ?? TextStyle;

            AddChildToLast(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyle = (CellStyle)itemStyle
            });
        }

        /// <summary>
        /// 增加一筆子DataColumn至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="dataKey">對應資料的Property</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleExpression">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void AddChildToLast<TProperty>(
            string headerText, string dataKey,
            Func<TProperty, T, object> render,
            CellStyle? headerStyle, Func<TProperty, T, CellStyle> dataStyleExpression
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            AddChildToLast(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = dataKey,
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleExpression
            });
        }

        /// <summary>
        /// 增加一筆子DataColumn至目前最後一個DataColumn底下
        /// </summary>
        /// <param name="headerText">標題要顯示的文字</param>
        /// <param name="expression">用Expression設定DataKey</param>
        /// <param name="render">修正顯示資料內容的委派，委派第一個參數為DatayKey對應的資料值，第二個參數為整筆資料物件</param>
        /// <param name="headerStyle">標題的儲存格格式，預設ListTemplateUtils.HeaderStyle</param>
        /// <param name="dataStyleExpression">資料的儲存格式，依資料決定顯示內容的委派</param>
        public void AddChildToLast<TProperty>(
            string headerText, Expression<Func<T, object>> expression,
            Func<TProperty, T, object> render,
            CellStyle? headerStyle, Func<TProperty, T, CellStyle> dataStyleExpression
        ) {
            headerStyle = headerStyle ?? HeaderStyle;

            AddChildToLast(new DataColumn<TProperty, T>() {
                HeaderText = headerText,
                DataKey = GetDataKeyByExpression(expression),
                ContentRender = render,
                HeaderStyle = (CellStyle)headerStyle,
                ItemStyleFunctor = dataStyleExpression
            });
        }

        protected override void InsertItem(int index, DataColumn<T> item) {
            if (item.ParentColumns != null) {
                throw new ArgumentException($"{nameof(DataColumn<T>)}重複加入!", nameof(item));
            }

            base.InsertItem(index, item);
            item.ParentColumns = this;
            ResetRootPoint();
        }

        protected override void RemoveItem(int index) {
            Items[index].ParentColumns = null;
            base.RemoveItem(index);
            ResetRootPoint();
        }

        protected override void SetItem(int index, DataColumn<T> item) {
            if (item.ParentColumns != null) {
                throw new ArgumentException($"{nameof(DataColumn<T>)}重複加入!", nameof(item));
            }

            Items[index].ParentColumns = null;
            base.SetItem(index, item);
            item.ParentColumns = this;
            ResetRootPoint();
        }

        protected override void ClearItems() {
            foreach (DataColumn<T> item in Items) {
                item.ParentColumns = null;
            }
            base.ClearItems();
            ResetRootPoint();
        }

        /// <summary>
        /// 和DataSource關聯的的DataColumns集合
        /// </summary>
        public IEnumerable<DataColumn<T>> DataSourceColumns {
            get {
                foreach (DataColumn<T> item in this) {
                    if (item.ChildColumns.Count == 0) {
                        yield return item;
                    } else {
                        foreach (DataColumn<T> dataHeader in item.ChildColumns.DataSourceColumns) {
                            yield return dataHeader;
                        }
                    }
                }
            }
        }
    }
}