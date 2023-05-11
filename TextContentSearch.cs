using System.Data;
using System.Management.Automation;
using OfficeOpenXml;


namespace TextContentSearch
{
    [Cmdlet(VerbsCommon.Search, "TextContent")]
    public class TextContentSearch : PSCmdlet
    {
        [Parameter(Mandatory = true)]
        public string? SearchKeyword { get; set; }

        [Parameter(Mandatory = true)]
        public string? DirectoryPath { get; set; }

        [Parameter(Mandatory = true)]
        public string? OutputPath { get; set; }

        protected override void ProcessRecord()
        {
            try
            {
                if (string.IsNullOrEmpty(SearchKeyword))
                {
                    throw new ArgumentNullException(nameof(SearchKeyword), "Search keyword cannot be null or empty.");
                }

                if (string.IsNullOrEmpty(DirectoryPath))
                {
                    throw new ArgumentNullException(nameof(DirectoryPath), "Directory path cannot be null or empty.");
                }

                var files = Directory.GetFiles(DirectoryPath);

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Search Results");

                    worksheet.Cells[1, 1].Value = "File Path";
                    worksheet.Cells[1, 2].Value = "Line Number";
                    worksheet.Cells[1, 3].Value = "Line Text";

                    var rowIndex = 2;

                    foreach (var file in files)
                    {
                        var lines = File.ReadLines(file);
                        var query = from line in lines.Select((text, index) => new { Text = text, LineNumber = index + 1 })
                                    where line.Text.Contains(SearchKeyword)
                                    select new { FilePath = file, LineNumber = line.LineNumber, LineText = line.Text };

                        foreach (var result in query)
                        {
                            worksheet.Cells[rowIndex, 1].Value = result.FilePath;
                            worksheet.Cells[rowIndex, 2].Value = result.LineNumber;
                            worksheet.Cells[rowIndex, 3].Value = result.LineText;

                            rowIndex++;
                        }
                    }

                    package.SaveAs(new FileInfo(OutputPath));
                }
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(ex, "1", ErrorCategory.InvalidArgument, null));
            }
        }
    }
}
