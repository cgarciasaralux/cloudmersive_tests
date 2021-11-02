using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentConverter.Lib
{
    public class FormCitation
    {
        public string FormType { get; set; }
        public string FormTitle { get; set; }
        public string BookTitle { get; set; }
        public List<string> Ref_chapters { get; set; }
        public string PermaLink { get; set; }
        public string AuthorString { get; set; }
        public string PlusLogoUrl { get; set; }

    }
}
