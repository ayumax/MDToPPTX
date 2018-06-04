using System;

namespace MD2PPTX
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
