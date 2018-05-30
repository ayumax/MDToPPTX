using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXLink
    {
        public string LinkKey { get; set; } = "";
        public string LinkURL { get; set; } = "";

        public bool IsEnable => !string.IsNullOrWhiteSpace(LinkKey) && !string.IsNullOrWhiteSpace(LinkURL);
    }
}
