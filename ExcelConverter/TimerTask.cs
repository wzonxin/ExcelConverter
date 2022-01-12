using System;

namespace ExcelConverter
{
    public enum TaskType 
    {
        SearchError,
        UpdateSearchProgress,
        FinishedSearch,
        NodeCheckedChanged,
        ConvertOutput,
        ConvertFinishWithFailed,
    }

    interface ITask
    {
        void DoAction();
    }

    class TimerTask : ITask
    {
        public Action Action;

        public void DoAction()
        {
            Action();
        }
    }
    
    class TimerTask<T> : ITask
    {
        public T Data;
        public Action<T> Action;

        public void DoAction()
        {
            Action(Data);
        }
    }

    class TimerTask<T, K> : ITask
    {
        public T Data;
        public K Data2;
        public Action<T, K> Action;

        public void DoAction()
        {
            Action(Data, Data2);
        }
    }
}
