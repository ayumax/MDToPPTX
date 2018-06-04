using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarkPP
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0) return;

            MDToPPTX.MDToPPTX pptxConverter = new MDToPPTX.MDToPPTX();

            string filepath = args[0];

            pptxConverter.Run(filepath);
        }
    }
}
