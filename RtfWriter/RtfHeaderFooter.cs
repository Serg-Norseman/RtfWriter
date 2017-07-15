using System;
using System.Collections;
using System.Text;

namespace Elistia.DotNetRtfWriter
{
    /// <summary>
    /// Summary description for RtfHeaderFooter
    /// </summary>
    public class RtfHeaderFooter : RtfBlockList
    {
        private Hashtable _magicWords;
        private HeaderFooterType _type;
        
        internal RtfHeaderFooter(HeaderFooterType type)
            : base(true, false, true, true, false)
        {
            _magicWords = new Hashtable();
            _type = type;
        }

        public override string render()
        {
            StringBuilder result = new StringBuilder();

            if (_type == HeaderFooterType.Header) {
                result.AppendLine(@"{\header");
            } else if (_type == HeaderFooterType.Footer) {
                result.AppendLine(@"{\footer");
            } else {
                throw new Exception("Invalid HeaderFooterType");
            }
            result.AppendLine();
            for (int i = 0; i < base.Blocks.Count; i++) {
                RtfBlock block = base.Blocks[i];
                if (base.DefaultCharFormat != null && block.DefaultCharFormat != null) {
                    block.DefaultCharFormat.copyFrom(base.DefaultCharFormat);
                }
                result.AppendLine(((RtfBlock)Blocks[i]).render());
            }
            result.AppendLine("}");
            return result.ToString();
        }
    }
}
