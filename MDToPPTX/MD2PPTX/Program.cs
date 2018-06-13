using System;

namespace MD2PPTX
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0) return;

            MDToPPTX.MD2PPTX pptxConverter = new MDToPPTX.MD2PPTX();

            string filepath = args[0];

            pptxConverter.Run(filepath);
        }
    }
}
