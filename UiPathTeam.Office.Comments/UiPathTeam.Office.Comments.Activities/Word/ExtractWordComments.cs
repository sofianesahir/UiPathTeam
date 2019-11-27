using Microsoft.Office.Interop.Word;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UiPathTeam.Office.Comments.Activities.Properties;
using DataTable = System.Data.DataTable;

namespace UiPathTeam.Office.Comments.Activities.Word
{
    public class ExtractWordComments : CodeActivity
    {
        [LocalizedCategory(nameof(Resources.Input))]
        public InArgument<String> FilePath { get; set; }

        [LocalizedCategory(nameof(Resources.Output))]
        public OutArgument<DataTable> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Application application = new Application();
            Document document = application.Documents.Open(FilePath.Get(context));
            DataRow row;

            DataTable output = new DataTable("output");

            DataColumn columnDate = new DataColumn("Date", System.Type.GetType("System.DateTime"));
            DataColumn columnAuthor = new DataColumn("Author", System.Type.GetType("System.String"));
            DataColumn columnComment = new DataColumn("Comment", System.Type.GetType("System.String"));


            output.Columns.Add(columnDate);
            output.Columns.Add(columnAuthor);
            output.Columns.Add(columnComment);


            foreach (Comment comment in document.Comments)
            {
                row = output.NewRow();
                row["Date"] = comment.Date;
                row["Author"] = comment.Author.ToString();
                row["Comment"] = comment.Scope.Text;
                output.Rows.Add(row);
            }

            application.Quit();
            Result.Set(context, output);

        }
    }
}
