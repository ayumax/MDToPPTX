using System;
using MDToPPTX;

namespace MD2PPTX
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0) return;

            MD2PPTX pptxConverter = new MD2PPTX();

            string filepath = args[0];
            string title = args.Length > 1 ? args[1] : "";
            string subTitle = args.Length > 2 ? args[2] : "";

            MDToPPTX.PPTX.PPTXSetting setting = new MDToPPTX.PPTX.PPTXSetting()
            {
                SlideSize = MDToPPTX.PPTX.EPPTXSlideSizeValues.Screen4x3,
                Title = title,
                SubTitle = subTitle
            };

            pptxConverter.Run(filepath, setting);
        }
    }
}
