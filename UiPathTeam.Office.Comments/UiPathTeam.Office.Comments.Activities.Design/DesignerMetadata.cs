using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using UiPathTeam.Office.Comments.Activities.Design.Designers;
using UiPathTeam.Office.Comments.Activities.Design.Properties;
using UiPathTeam.Office.Comments.Activities.Excel;
using UiPathTeam.Office.Comments.Activities.Word;

namespace UiPathTeam.Office.Comments.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute =  new CategoryAttribute($"{Resources.Category}");


            builder.AddCustomAttributes(typeof(ExtractWordComments), categoryAttribute);
            builder.AddCustomAttributes(typeof(ExtractWordComments), new DesignerAttribute(typeof(ExtractWordCommentsDesigner)));
            builder.AddCustomAttributes(typeof(ExtractWordComments), new HelpKeywordAttribute("https://go.uipath.com"));

            builder.AddCustomAttributes(typeof(ExtractExcelComments), categoryAttribute);
            builder.AddCustomAttributes(typeof(ExtractExcelComments), new DesignerAttribute(typeof(ExtractExcelCommentsDesigner)));
            builder.AddCustomAttributes(typeof(ExtractExcelComments), new HelpKeywordAttribute("https://go.uipath.com"));

            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
