using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace FormPdf.AddIn.Excel
{
    /// <summary>
    /// 中間の COM オブジェクト（Range, PageSetup, TextFrame2 ...）をまとめて解放するためのスコープ。
    /// <code>using (var com = new ComScope()) { var r = com.Add(ws.UsedRange); ... }</code>
    /// </summary>
    internal sealed class ComScope : IDisposable
    {
        private readonly List<object> _items = new List<object>();

        public T Add<T>(T obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
                _items.Add(obj);
            return obj;
        }

        public void Dispose()
        {
            for (int i = _items.Count - 1; i >= 0; i--)
                Release(_items[i]);
            _items.Clear();
        }

        public static void Release(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // 解放失敗は無視
            }
        }
    }
}
