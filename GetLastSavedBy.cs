using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
namespace GetLastUpdatedBy
{
    public class GetLastSavedBy:CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Output")]
        public OutArgument<string> LastUpdatedBy { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                var filePath = FilePath.Get(context);
                string lastSavedBy = null;
                using (var so = ShellObject.FromParsingName(filePath))
                {
                    var lastAuthorProperty = so.Properties.GetProperty(SystemProperties.System.Document.LastAuthor);
                    if (lastAuthorProperty != null)
                    {
                        var lastAuthor = lastAuthorProperty.ValueAsObject;
                        if (lastAuthor != null)
                        {
                            lastSavedBy = lastAuthor.ToString();
                        }
                    }
                }
                LastUpdatedBy.Set(context, lastSavedBy);
            }
            catch (Exception ex)
            {

            }

        }
    }
}
