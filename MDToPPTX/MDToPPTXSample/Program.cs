using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDToPPTXSample
{
    class Program
    {
        static void Main(string[] args)
        {
            MDToPPTX.MDToPPTX pptxConverter = new MDToPPTX.MDToPPTX();

            string filepath = @"C:\Users\ayuma\Desktop\sample3.pptx";

            pptxConverter.Run(filepath);
        }
    }
}
