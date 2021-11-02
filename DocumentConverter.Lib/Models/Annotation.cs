using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentConverter.Lib.Models
{
    public class Annotation
    {
        public string Title { get; set; }
        public string AnnotationType { get; set; }
        public int? PageNumber { get; set; }
        public int? AnnotationIndex { get; set; }
        public string Subject { get; set; }
        public string TextContents { get; set; }
        public DateTime? CreationDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public double? LeftX { get; set; }
        public double? TopY { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }
    }
}
