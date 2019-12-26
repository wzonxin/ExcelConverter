using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConverter
{
    public enum TaskType
    {
        UpdateProgress,
        FinishSearch,
    }

    class TimerTask
    {
        public object Data;
        public TaskType TaskType;
        public Action<object> Action;
    }
}
