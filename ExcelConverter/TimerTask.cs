using System;

namespace ExcelConverter
{
    public enum TaskType 
    {
        SearchError,
        UpdateSearchProgress,
        FinshedSearch,
    }

    interface ITask
    {
        void DoAction();
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
