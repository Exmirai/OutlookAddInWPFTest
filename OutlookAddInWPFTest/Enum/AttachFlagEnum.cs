using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInWPFTest.Enum
{
    [Flags]
    public enum AttachFlagEnum
    {
        LEFT = 0,
        RIGHT = 1,
        UP = 2,
        DOWN = 4,
        INSIDE = 8,
        OUTSIDE = 16,
        OVERLAY = 32,
    }
}
