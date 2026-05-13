using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIn.Infrastructure.Threading
{
    internal sealed class OutlookThreadInvoker
    {
        private readonly Control _control;

        public OutlookThreadInvoker(Control control)
        {
            _control = control ?? throw new ArgumentNullException(nameof(control));
        }

        public Task<T> InvokeAsync<T>(Func<T> work)
        {
            if (_control.IsDisposed || !_control.IsHandleCreated)
                throw new InvalidOperationException("Outlook UI context is not ready.");

            if (!_control.InvokeRequired)
                return Task.FromResult(work());

            var tcs = new TaskCompletionSource<T>();
            _control.BeginInvoke((Action)(() =>
            {
                try { tcs.TrySetResult(work()); }
                catch (Exception ex) { tcs.TrySetException(ex); }
            }));
            return tcs.Task;
        }

        public Task InvokeAsync(Action work)
        {
            return InvokeAsync<object>(() =>
            {
                work();
                return null;
            });
        }

        public Task InvokeLegacyAsyncCommand(Func<Task> work)
        {
            if (_control.IsDisposed || !_control.IsHandleCreated)
                throw new InvalidOperationException("Outlook UI context is not ready.");

            if (!_control.InvokeRequired)
                return work();

            var tcs = new TaskCompletionSource<object>();
            _control.BeginInvoke((Action)(async () =>
            {
                try
                {
                    await work();
                    tcs.TrySetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }));
            return tcs.Task;
        }
    }
}
