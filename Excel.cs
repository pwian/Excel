using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Excel
{
    /// <summary>
    /// Класс для работы с excel файлами.
    /// </summary>
    public sealed class Excel : IDisposable
    {
        private Workbook _workbook;
        private Worksheet _worksheet;

        /// <summary>
        /// Конструктор.
        /// </summary>
        /// <param name="inPath">Путь к файлу.</param>
        public Excel(string inPath)
        {
            _workbook = new Application().Workbooks.Open(inPath);
            _worksheet = _workbook.Worksheets[1];
        }

        /// <summary>
        /// Возвращает содержимое ячейки.
        /// </summary>
        /// <param name="inRowIndex">Номер строки.</param>
        /// <param name="inColumnIndex">Номер колонки.</param>
        public string ReadCell(int inRowIndex, int inColumnIndex)
        {
            return Convert.ToString(_worksheet.Cells[inRowIndex, inColumnIndex].Value2);
        }

        /// <summary>
        /// Заполняет ячейку.
        /// </summary>
        /// <param name="inRowIndex">Номер строки.</param>
        /// <param name="inColumnIndex">Номер колонки.</param>
        /// <param name="inValue">Значение, которое нужно добавить в ячейку.</param>
        public void WriteToCell(int inRowIndex, int inColumnIndex, string inValue)
        {
            _worksheet.Cells[inRowIndex, inColumnIndex].Value2 = inValue;
        }

        /// <summary>
        /// Сохраняет файл.
        /// </summary>
        public void Save()
        {
            _workbook.Save();
        }

        /// <summary>
        /// Закрытие файла
        /// </summary>
        public void Close()
        {
            _workbook.Close();
        }

        /// <inheritdoc />
        public void Dispose()
        {
            if (_worksheet != null)
            {
                Marshal.FinalReleaseComObject(_worksheet);
                _worksheet = null;
            }

            if (_workbook != null)
            {
                Marshal.FinalReleaseComObject(_workbook);
                _workbook = null;
            }
        }
    }

}
