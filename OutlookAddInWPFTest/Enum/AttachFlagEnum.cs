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
        DOWN = 3,
        INSIDE = 4,
        OUTSIDE = 5,
        OVERLAY = 6,
    }
}
