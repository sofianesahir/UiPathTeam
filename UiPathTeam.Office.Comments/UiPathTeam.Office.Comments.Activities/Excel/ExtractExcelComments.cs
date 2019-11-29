using Microsoft.Office.Interop.Excel;
using System;
using System.Activities;
using System.Data;
using System.IO;
using UiPathTeam.Office.Comments.Activities.Properties;
using DataTable = System.Data.DataTable;

namespace UiPathTeam.Office.Comments.Activities.Excel
{
    public class ExtractExcelComments : CodeActivity
    {
        [LocalizedCategory(nameof(Resources.Input))]
        [LocalizedDescription(nameof(Resources.FilePathDescription))]
        [LocalizedDisplayName(nameof(Resources.FilePathDisplayName))]
        public InArgument<String> FilePath { get; set; }

        [LocalizedCategory(nameof(Resources.Options))]
        [LocalizedDescription(nameof(Resources.ExtractDateDescription))]
        [LocalizedDisplayName(nameof(Resources.ExtractDateDisplayName))]
        public bool ExtractDate { get; set; }

        [LocalizedCategory(nameof(Resources.Options))]
        [LocalizedDescription(nameof(Resources.ExtractAuthorDescription))]
        [LocalizedDisplayName(nameof(Resources.ExtractAuthorDisplayName))]
        public bool ExtractAuthor { get; set; }

        [LocalizedCategory(nameof(Resources.Options))]
        [LocalizedDescription(nameof(Resources.ExtractCommentDescription))]
        [LocalizedDisplayName(nameof(Resources.ExtractCommentDisplayName))]
        public bool ExtractComment { get; set; }

        [LocalizedCategory(nameof(Resources.Output))]
        [LocalizedDescription(nameof(Resources.ResultDescription))]
        [LocalizedDisplayName(nameof(Resources.ResultDisplayName))]
        public OutArgument<DataTable> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Application application = new Application();
            Workbook workbook = application.Workbooks.Open(Path.GetFullPath(FilePath.Get(context)));
            Worksheet worksheet = workbook.Sheets[1];

            DataRow row;

            DataTable output = new DataTable("output");

            DataColumn columnDate = new DataColumn("Date", System.Type.GetType("System.DateTime"));
            DataColumn columnAuthor = new DataColumn("Author", System.Type.GetType("System.String"));
            DataColumn columnComment = new DataColumn("Comment", System.Type.GetType("System.String"));

            if (ExtractDate)
                output.Columns.Add(columnDate);
            if (ExtractAuthor)
                output.Columns.Add(columnAuthor);
            if (ExtractComment)
                output.Columns.Add(columnComment);


            foreach (Comment comment in worksheet.Comments)
            {

                row = output.NewRow();
                if (ExtractDate)
                    row["Date"] = DateTime.Now;
                if (ExtractAuthor)
                    row["Author"] = comment.Author.ToString();
                if (ExtractComment)
                    row["Comment"] = comment.Text();

                output.Rows.Add(row);
            }

            workbook.Close();
            application.Quit();
            Result.Set(context, output);

        }
    }
}
