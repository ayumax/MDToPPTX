using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class Utils
    {
        private const int cm2shapescale = 360000;
        private const int degree2shapescale = 60000;

        public static long GetCmToShapeScale(float CmValue)
        {
            return (long)(CmValue * cm2shapescale);
        }

        public static int GetDegreeToShapeScale(int DegreeValue)
        {
            return DegreeValue * degree2shapescale;
        }
    }
}
