using System;
using System.Collections.Generic;

namespace AccountingRobot
{
    public interface ICachedValue<T>
    {
        List<T> GetLatestValues(bool forceUpdate = false);
    }
}
