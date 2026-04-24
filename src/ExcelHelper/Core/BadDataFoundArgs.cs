namespace ExcelHelper.Core
{
    /// <summary>
    /// Arguments provided to the <see cref="ExcelConfiguration.BadDataFound"/> callback.
    /// </summary>
    public sealed class BadDataFoundArgs
    {
        /// <summary>
        /// Gets the name of the field where bad data was found, if available.
        /// </summary>
        public string? Field { get; }

        /// <summary>
        /// Gets the raw cell value.
        /// </summary>
        public string? RawCellValue { get; }

        /// <summary>
        /// Gets the 1-based row index.
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Gets the 1-based column index.
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="BadDataFoundArgs"/> class.
        /// </summary>
        public BadDataFoundArgs(string? field, string? rawCellValue, int row, int column)
        {
            Field = field;
            RawCellValue = rawCellValue;
            Row = row;
            Column = column;
        }
    }
}
