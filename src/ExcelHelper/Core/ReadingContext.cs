namespace ExcelHelper.Core
{
    /// <summary>
    /// Provides context information during reading operations.
    /// </summary>
    public sealed class ReadingContext : ExcelContext
    {
        /// <summary>
        /// Gets the header record if <see cref="ExcelConfiguration.HasHeaderRecord"/> is true.
        /// </summary>
        public string[]? HeaderRecord { get; internal set; }

        /// <summary>
        /// Gets the current raw record values.
        /// </summary>
        public string[]? Record { get; internal set; }

        /// <summary>
        /// Gets the current 0-based record index within the data rows.
        /// </summary>
        public int CurrentIndex { get; internal set; }

        /// <summary>
        /// Gets the total number of rows in the worksheet.
        /// </summary>
        public int RowCount { get; internal set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ReadingContext"/> class.
        /// </summary>
        /// <param name="configuration">The configuration for the reading operation.</param>
        public ReadingContext(ExcelConfiguration configuration)
            : base(configuration)
        {
        }
    }
}
