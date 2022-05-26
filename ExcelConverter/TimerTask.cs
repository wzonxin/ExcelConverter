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

    abstract class TaskBase : IMemPoolItem
    {
        public abstract void DoAction();
        public abstract void ClearData();
        public abstract void Recycle();

        public sealed override void OnDestroy()
        {
            ClearData();
        }
    }

    class TimerTask : TaskBase
    {
        public Action Action;

        public override void DoAction()
        {
            Action();
        }

        public override void ClearData()
        {
            Action = null;
        }

        public override void Recycle()
        {
            MemPool<TimerTask>.Recycle(this);
        }
    }
    
    class TimerTask<T> : TaskBase
    {
        public T Data;
        public Action<T> Action;

        public override void DoAction()
        {
            Action(Data);
        }

        public override void ClearData()
        {
            Data = default(T);
            Action = null;
        }

        public override void Recycle()
        {
            MemPool<TimerTask<T>>.Recycle(this);
        }
    }

    class TimerTask<T, K> : TaskBase
    {
        public T Data;
        public K Data2;
        public Action<T, K> Action;

        public override void DoAction()
        {
            Action(Data, Data2);
        }

        public override void Recycle()
        {
            MemPool<TimerTask<T, K>>.Recycle(this);
        }

        public override void ClearData()
        {
            Data = default(T);
            Data2 = default(K);
            Action = null;
        }
    }
}
