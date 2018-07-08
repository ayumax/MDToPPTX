using System;
using MDToPPTX;
using MDToPPTX.PPTX;

namespace MarkPP
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
            string settingPath = args.Length > 3 ? args[3] : "";

            PPTXSetting setting = null;

            if (string.IsNullOrWhiteSpace(settingPath) == false)
            {
                if (System.IO.File.Exists(settingPath))
                {
                    setting = PPTXSetting.Load(settingPath);
                }
            }

            setting = setting ?? new PPTXSetting()
            {
                SlideSize = EPPTXSlideSizeValues.Screen4x3,
                Title = title,
                SubTitle = subTitle
            };

            pptxConverter.RunFromMDFile(filepath, null, setting);

            //setting.Save(settingPath);
        }
    }
}
