using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using ExcelDna.Integration;
using Timer = System.Timers.Timer;
using Newtonsoft.Json;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace AsyncBatching
{
    public class Stock
    {
        public string CompanyName { get; set; }
        public string CompanyCode { get; set; }
        public float MarketPrice { get; set; }
        public string Volume { get; set; }
    }

    public static class JsonFunctions
    {
        static string address = AppDomain.CurrentDomain.BaseDirectory+ @"\Stocks.json";

        public static object CompanyNameSync(string code)
        {
            var list = JsonConvert.DeserializeObject<List<Stock>>(File.ReadAllText(@address));
            List<string> CompanyList = new List<string>();

            foreach (var item in list)
            {
                if (item.CompanyCode.Contains(code))
                    CompanyList.Add(item.CompanyName);
            }

            if (CompanyList.Count == 0)
                return "Company not found !";
            else
                return string.Join(",", CompanyList.ToArray());
        }

        public static object CompanyName(string code)
        {
            for(int i=0; i !=100000; i++)
            {
                for (int j = 0; j != 1000; j++) ;
            };

            var list = JsonConvert.DeserializeObject<List<Stock>>(File.ReadAllText(@address));
            List<string> CompanyList = new List<string>();

            foreach (var item in list)
            {
                if (item.CompanyCode.Contains(code))
                    CompanyList.Add(item.CompanyName);
            }

            if (CompanyList.Count == 0)
                return "Company not found !";
            else
                return string.Join(",", CompanyList.ToArray());
        }

        public static object StockPrice(string name)
        {
            var list = JsonConvert.DeserializeObject<List<Stock>>(File.ReadAllText(@address));

            foreach (var item in list)
            {
                if (item.CompanyName == name || item.CompanyCode == name)
                    return item.MarketPrice;
            }
            return "Company not found - Incorrect Company name or code !";
        }

        public static object Volume(string name)
        {
            var list = JsonConvert.DeserializeObject<List<Stock>>(File.ReadAllText(@address));

            foreach (var item in list)
            {
                if (item.CompanyName == name || item.CompanyCode == name)
                    return item.Volume;
            }
            return "Company not found - Incorrect Company name or code !";
        }

         public static string dataPuller()
        {
            var list = JsonConvert.DeserializeObject<List<Stock>>(File.ReadAllText(@address));
            List<string> CompanyList = new List<string>();

            foreach (var item in list)
            {
                CompanyList.Add(item.CompanyCode);
            }
            return string.Join(",", CompanyList.ToArray());
        } //Just to extract data from a JSON -  
                                              //NO USE AFTER SOLUTION IS DEPLOYED WITH DC
    }

    public static class AsyncBatchExample
    {
        static int MaxBatchSize = 10;
        static TimeSpan Timeout = TimeSpan.FromMilliseconds(250);

        // Instantiating an AsyncBatchUtil class object with parameters MAX_BATCH_SIZE, TIMEOUT, TASK<List<Object>>
        static readonly AsyncBatchUtil BatchRunner = new AsyncBatchUtil(MaxBatchSize, Timeout , RunBatch);

        // This function will be called for each batch, on a ThreadPool thread.
        // Each AsyncCall contains the function name and arguments passed from the function.
        // The List<object> returned by the Task must contain the results, corresponding to the calls list.

        static async Task<List<object>> RunBatch(List<AsyncBatchUtil.AsyncCall> calls)
        {
            var batchStart = DateTime.Now;

            // Simulate things taking a while...
            // await Task.Delay(TimeSpan.FromSeconds(5));
           
            await Task.Run(() =>
            {
                // Loop to generate delay
                for (int i = 0; i != 1000000; ++i)
                    for (int j = 0; j != 1000; j++);
            });


            // Building up the list of results...
            var results = new List<object>(5);
            foreach (var call in calls)
            {
                var result = string.Format("{0}", call.Arguments[0]);
                results.Add(result);
            }

            return results;

        }

        static public object GetCompanyName(string code)
        {
            return BatchRunner.Run("GetCompanyName", JsonFunctions.CompanyNameSync(code));
        } //Async method for CompanyName

        static public object GetStockPrice(string code)
        {
            return BatchRunner.Run("GetStockPrice", JsonFunctions.StockPrice(code));
        }   //Async method for StockPrice

        static public object GetMarketVolume(string code)
        {
            return BatchRunner.Run("GetmarketVolume", JsonFunctions.Volume(code));
        } //Async method for MarketVolume

    }

    // This is the main helper class for supporting batched async calls
    public class AsyncBatchUtil
    {
        // Represents a single function call in  a batch
        public class AsyncCall
        {
            internal TaskCompletionSource<object> TaskCompletionSource;
            public string FunctionName { get; private set; }
            public object[] Arguments { get; private set; }

            public AsyncCall(TaskCompletionSource<object> taskCompletion, string functionName, object[] args)
            {
                TaskCompletionSource = taskCompletion;
                FunctionName = functionName;
                Arguments = args;
            }
        }

        // Not a hard limit
        readonly int _maxBatchSize;
        readonly Func<List<AsyncCall>, Task<List<object>>> _batchRunner;

        readonly object _lock = new object();
        readonly Timer _batchTimer;   // Timer events will fire from a ThreadPool thread
        List<AsyncCall> _currentBatch;

        public AsyncBatchUtil(int maxBatchSize, TimeSpan batchTimeout, Func<List<AsyncCall>, Task<List<object>>> batchRunner)
        {
            if (maxBatchSize < 1)
            {
                throw new ArgumentOutOfRangeException("maxBatchSize", "Max batch size must be non zero and positive");
            }
            if (batchRunner == null)
            {
                // Check early - otherwise the NullReferenceException would happen in a threadpool callback.
                throw new ArgumentNullException("batchRunner");
            }

            _maxBatchSize = maxBatchSize;
            _batchRunner = batchRunner;

            _currentBatch = new List<AsyncCall>();

            _batchTimer = new Timer(batchTimeout.TotalMilliseconds);
            _batchTimer.AutoReset = false;
            _batchTimer.Elapsed += TimerElapsed;
            // Timer is not Enabled (Started) by default
        }

        // Will only run on the main thread
        public object Run(string functionName, params object[] args)
        {
            return ExcelAsyncUtil.Observe(functionName, args, delegate
            {
                var tcs = new TaskCompletionSource<object>();
                EnqueueAsyncCall(tcs, functionName, args);
                return new TaskExcelObservable(tcs.Task);
            });
        }

        // Will only run on the main thread
        void EnqueueAsyncCall(TaskCompletionSource<object> taskCompletion, string functionName, object[] args)
        {
            lock (_lock)
            {
                _currentBatch.Add(new AsyncCall(taskCompletion, functionName, args));

                // Check if the batch size has been reached, schedule it to be run
                if (_currentBatch.Count >= _maxBatchSize)
                {
                    // This won't run the batch immediately, but will ensure that the current batch (containing this call) will run soon.
                    ThreadPool.QueueUserWorkItem(state => RunBatch((List<AsyncCall>)state), _currentBatch);
                    _currentBatch = new List<AsyncCall>();
                    _batchTimer.Stop();
                }
                else
                {
                    // We don't know if the batch containing the current call will run, 
                    // so ensure that a timer is started.
                    if (!_batchTimer.Enabled)
                    {
                        _batchTimer.Start();
                    }
                }
            }
        }

        // Will run on a ThreadPool thread
        void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            List<AsyncCall> batch;
            lock (_lock)
            {
                batch = _currentBatch;
                _currentBatch = new List<AsyncCall>();
            }
            RunBatch(batch);
        }


        // Will always run on a ThreadPool thread
        // Might be re-entered...
        // batch is allowed to be empty
        async void RunBatch(List<AsyncCall> batch)
        {
            // Maybe due to Timer re-entrancy we got an empty batch...?
            if (batch.Count == 0)
            {
                // No problem - just return
                return;
            }

            try
            {
                var resultList = await _batchRunner(batch);
                if (resultList.Count != batch.Count)
                {
                    throw new InvalidOperationException(string.Format("Batch result size incorrect. Batch Count: {0}, Result Count: {1}", batch.Count, resultList.Count));
                }

                for (int i = 0; i < resultList.Count; i++)
                {
                    batch[i].TaskCompletionSource.SetResult(resultList[i]);
                }
            }
            catch (Exception ex)
            {
                foreach (var call in batch)
                {
                    call.TaskCompletionSource.SetException(ex);
                }
            }
        }

        // Helper class to turn a task into an IExcelObservable that either returns the task result and completes, or pushes an Exception
        class TaskExcelObservable : IExcelObservable
        {
            readonly Task<object> _task;

            public TaskExcelObservable(Task<object> task)
            {
                _task = task;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                switch (_task.Status)
                {
                    case TaskStatus.RanToCompletion:
                        observer.OnNext(_task.Result);
                        observer.OnCompleted();
                        break;
                    case TaskStatus.Faulted:
                        observer.OnError(_task.Exception.InnerException);
                        break;
                    case TaskStatus.Canceled:
                        observer.OnError(new TaskCanceledException(_task));
                        break;
                    default:
                        var task = _task;
                        // OK - the Task has not completed synchronously
                        // And handle the Task completion
                        task.ContinueWith(t =>
                        {
                            switch (t.Status)
                            {
                                case TaskStatus.RanToCompletion:
                                    observer.OnNext(t.Result);
                                    observer.OnCompleted();
                                    break;
                                case TaskStatus.Faulted:
                                    observer.OnError(t.Exception.InnerException);
                                    break;
                                case TaskStatus.Canceled:
                                    observer.OnError(new TaskCanceledException(t));
                                    break;
                            }
                        });
                        break;
                }

                return DefaultDisposable.Instance;
            }

            // Helper class to make an empty IDisposable
            sealed class DefaultDisposable : IDisposable
            {
                public static readonly DefaultDisposable Instance = new DefaultDisposable();
                // Prevent external instantiation
                DefaultDisposable() { }
                public void Dispose() { }
            }
        }
    }
}