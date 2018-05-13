using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public enum PPTXBullet
    {
        /// <summary>
        /// 箇条書きなし
        /// </summary>
        None,
        /// <summary>
        /// ●
        /// </summary>
        Circle,
        /// <summary>
        /// ■
        /// </summary>
        Rectangle,
        /// <summary>
        /// ◆
        /// </summary>
        Diamond,
        /// <summary>
        /// □
        /// </summary>
        RectangleBorder,
        /// <summary>
        /// ✔
        /// </summary>
        Check,
        /// <summary>
        /// ▶
        /// </summary>
        Arrow,
        /// <summary>
        /// ●（小さい）
        /// </summary>
        MiniCircle,
        /// <summary>
        /// 数字 1. 2. 3. 
        /// </summary>
        Number,
        /// <summary>
        /// 数字 ①　②　③
        /// </summary>
        CircleNumber
    }
}
