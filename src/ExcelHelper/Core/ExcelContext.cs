using OfficeOpenXml;
using System;

namespace ExcelHelper.Core
{
    /// <summary>
    /// Provides context information during Excel read and write operations.
    /// </summary>
    public abstract class ExcelContext
    {
        /// <summary>
        /// Gets the configuration used for the current operation.
        /// </summary>
        public ExcelConfiguration Configuration { get; }

        /// <summary>
        /// Gets the worksheet currently being processed.
        /// </summary>
        public ExcelWorksheet? Worksheet { get; internal set; }

        /// <summary>
        /// Gets the 1-based current row index.
        /// </summary>
        public int Row { get; internal set; }

        /// <summary>
        /// Gets the 1-based current column index.
        /// </summary>
        public int Column { get; internal set; }

        /// <summary>
        /// Gets the name of the current worksheet, or null if not set.
        /// </summary>
        public string? SheetName => Worksheet?.Name;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelContext"/> class.
        /// </summary>
        /// <param name="configuration">The configuration for the operation.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="configuration"/> is null.</exception>
        protected ExcelContext(ExcelConfiguration configuration)
        {
            Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }
    }
}
