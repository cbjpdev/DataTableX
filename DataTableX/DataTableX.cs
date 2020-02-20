using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;

namespace DatatableX
{
    /// <summary>
    /// Class for datatable extension methods
    /// </summary>
    public static class DataTableX
    {
        /// <summary>
        /// Convert list to data table
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns>Datatable</returns>
        public static DataTable ToDataTable<T>(this IList<T> list)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();

            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            object[] values = new object[props.Count];

            foreach (T item in list)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item) ?? DBNull.Value;
                }

                table.Rows.Add(values);
            }

            return table;
        }

        /// <summary>
        /// Convert IEnumerable to Data table
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable<T>(this IEnumerable<T> list)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();

            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            object[] values = new object[props.Count];

            foreach (T item in list)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item) ?? DBNull.Value;
                }

                table.Rows.Add(values);
            }

            return table;
        }

        /// <summary>
        /// List of dictionary type to datatable
        /// </summary>
        /// <param name="list"></param>
        /// <param name="tableData"></param>
        /// <param name="columnList"></param>
        /// <returns>Datatable</returns>
        public static DataTable ToDataTable(List<Dictionary<string, object>> list, Dictionary<string, Type> tableData, List<string> columnList)
        {
            DataTable result = new DataTable();

            if (list.Count == 0)
            {
                result.Columns.AddRange(columnList.Select(c => new DataColumn(c)).ToArray());
                result.Rows.Add(result.NewRow());
                return result;
            }

            var columnNames = tableData.Keys.ToList();

            result.Columns.AddRange(columnNames.Select(c =>
            {
                DataColumn dc = new DataColumn(c, tableData.GetValueOrDefault(c, typeof(string)))
                {
                    AllowDBNull = true
                };

                return dc;
            }).ToArray());

            foreach (Dictionary<string, object> item in list)
            {
                var row = result.NewRow();

                foreach (var key in item.Keys)
                {
                    row[key] = item[key] ?? DBNull.Value;
                }

                result.Rows.Add(row);
            }

            if (columnList != null)
            {
                for (int i = 0; i < columnList.Count; i++)
                {
                    result.Columns[i].ColumnName = columnList[i];
                    result.Columns[i].AllowDBNull = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Transform list to Datatable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static DataTable ToTransformeDataTable<T>(this IList<T> list)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();

            for (int i = 0; i < list.Count; i++)
            {
                string colName = list[i].GetType().GetProperties()[0].GetValue(list[i])?.ToString();
                table.Columns.Add(colName, typeof(double)); //// only for double
            }

            var maxCount = props.Count - 1; //// ignore first property

            for (int i = 0; i < maxCount; i++)
            {
                var row = table.NewRow();

                int j = 0;

                foreach (var item in list)
                {
                    //// [i + 1] starting from 1 to ignore first column
                    row[j] = item.GetType().GetProperties()[i + 1].GetValue(item) ?? DBNull.Value;
                    j++;
                }

                table.Rows.Add(row);
            }

            return table;
        } 
    }
}
